import React, { useEffect, useMemo, useState } from "react";
import {
  clearAllTasks,
  deleteSelectedTasks,
  deleteTask,
  getConfig,
  getPublicIp,
  getRpaLogTail,
  getRpaStatus,
  getStatus,
  listTasks,
  scan,
  sendBatch,
  sendSelected,
  startRpa,
  stopRpa,
  toggleAutoWatch,
  type RpaStartPayload,
  type RpaStatus,
  type Task,
  uploadContacts,
  updateConfig
} from "./api";
import TaskTable from "./components/TaskTable";

type TabKey = "api" | "rpa";

const defaultRpaForm: RpaStartPayload = {
  tasks_csv: "tools/rpa_tasks.real.csv",
  wecom_exe: "",
  main_title_re: ".*(WeCom|WXWork|企业微信).*",
  send_mode: "clipboard",
  dry_run: false,
  paste_only: true,
  no_chat_verify: true,
  interval_sec: 3,
  timeout_sec: 20,
  max_retries: 2,
  retry_delay_sec: 1.0,
  stabilize_open_rounds: 2,
  stabilize_focus_rounds: 2,
  stabilize_send_rounds: 1,
  open_chat_strategy: "keyboard_first",
  resume_from: 1,
  resume_failed: false,
  stop_on_fail: true,
  skip_missing_image: true,
  debug_chat_text: false,
  results_csv: "run-logs/rpa-results.csv",
  log_file: "run-logs/rpa-sender.log"
};

const emptyRpaStatus: RpaStatus = {
  running: false,
  pid: null,
  started_at: null,
  finished_at: null,
  return_code: null,
  command: [],
  log_file: "",
  results_csv: "",
  result_counts: {}
};

export default function App() {
  const [activeTab, setActiveTab] = useState<TabKey>("rpa");

  const [tasks, setTasks] = useState<Task[]>([]);
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [rootPath, setRootPath] = useState("/data/inbox");
  const [autoWatch, setAutoWatch] = useState(false);
  const [status, setStatus] = useState<any>(null);
  const [config, setConfig] = useState<any>(null);
  const [corpId, setCorpId] = useState("");
  const [agentId, setAgentId] = useState("");
  const [secret, setSecret] = useState("");
  const [publicIp, setPublicIp] = useState<string | null>(null);
  const [message, setMessage] = useState<string | null>(null);

  const [rpaForm, setRpaForm] = useState<RpaStartPayload>(defaultRpaForm);
  const [rpaStatus, setRpaStatus] = useState<RpaStatus>(emptyRpaStatus);
  const [rpaLogs, setRpaLogs] = useState<string[]>([]);
  const [rpaMessage, setRpaMessage] = useState<string | null>(null);
  const [rpaLoading, setRpaLoading] = useState(false);

  const selectedCount = useMemo(() => selected.size, [selected]);
  const rpaCounts = rpaStatus.result_counts || {};

  const rpaCmdPreview = useMemo(() => {
    const cmd = [
      "python",
      "tools/wecom_rpa_sender.py",
      "--tasks-csv",
      rpaForm.tasks_csv,
      "--send-mode",
      rpaForm.send_mode,
      "--main-title-re",
      rpaForm.main_title_re,
      "--interval-sec",
      String(rpaForm.interval_sec),
      "--timeout-sec",
      String(rpaForm.timeout_sec),
      "--max-retries",
      String(rpaForm.max_retries),
      "--retry-delay-sec",
      String(rpaForm.retry_delay_sec),
      "--stabilize-open-rounds",
      String(rpaForm.stabilize_open_rounds),
      "--stabilize-focus-rounds",
      String(rpaForm.stabilize_focus_rounds),
      "--stabilize-send-rounds",
      String(rpaForm.stabilize_send_rounds),
      "--open-chat-strategy",
      rpaForm.open_chat_strategy,
      "--resume-from",
      String(rpaForm.resume_from),
      "--results-csv",
      rpaForm.results_csv,
      "--log-file",
      rpaForm.log_file
    ];
    if (rpaForm.wecom_exe?.trim()) cmd.push("--wecom-exe", rpaForm.wecom_exe.trim());
    if (rpaForm.dry_run) cmd.push("--dry-run");
    if (rpaForm.paste_only) cmd.push("--paste-only");
    if (rpaForm.no_chat_verify) cmd.push("--no-chat-verify");
    if (rpaForm.resume_failed) cmd.push("--resume-failed", "--resume-results-csv", rpaForm.results_csv);
    if (rpaForm.stop_on_fail) cmd.push("--stop-on-fail");
    if (rpaForm.skip_missing_image) cmd.push("--skip-missing-image");
    if (rpaForm.debug_chat_text) cmd.push("--debug-chat-text");
    return cmd.map((s) => (s.includes(" ") ? `"${s}"` : s)).join(" ");
  }, [rpaForm]);

  async function refreshAll() {
    const [t, s, c] = await Promise.all([listTasks(), getStatus(), getConfig()]);
    setTasks(t);
    setStatus(s);
    setConfig(c);
    setRootPath(c.root_path || "/data/inbox");
  }

  async function refreshRpa() {
    const [s, log] = await Promise.all([getRpaStatus(), getRpaLogTail(180)]);
    setRpaStatus(s);
    setRpaLogs(log.lines || []);
  }

  useEffect(() => {
    refreshAll().catch((e) => setMessage(e.message));
    refreshRpa().catch(() => undefined);

    const id = setInterval(() => {
      refreshAll().catch(() => undefined);
      if (activeTab === "rpa" || rpaStatus.running) {
        refreshRpa().catch(() => undefined);
      }
    }, 2500);
    return () => clearInterval(id);
  }, [activeTab, rpaStatus.running]);

  function toggleSelection(id: number) {
    const next = new Set(selected);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    setSelected(next);
  }

  async function onScan() {
    try {
      await scan(rootPath);
      await refreshAll();
      setMessage("扫描完成");
    } catch (e: any) {
      setMessage(e.message);
    }
  }

  async function onBatchSend() {
    try {
      await sendBatch();
      await refreshAll();
      setMessage("批量发送已开始");
    } catch (e: any) {
      setMessage(e.message);
    }
  }

  async function onSelectedSend() {
    try {
      await sendSelected(Array.from(selected));
      setSelected(new Set());
      await refreshAll();
      setMessage("选中发送已开始");
    } catch (e: any) {
      setMessage(e.message);
    }
  }

  async function onDeleteOne(taskId: number) {
    if (!window.confirm("确认删除这条记录吗？")) return;
    try {
      await deleteTask(taskId);
      const next = new Set(selected);
      next.delete(taskId);
      setSelected(next);
      await refreshAll();
      setMessage("已删除记录");
    } catch (e: any) {
      setMessage(e.message);
    }
  }

  async function onDeleteSelected() {
    const ids = Array.from(selected);
    if (ids.length === 0) return;
    if (!window.confirm(`确认删除选中的 ${ids.length} 条记录吗？`)) return;
    try {
      await deleteSelectedTasks(ids);
      setSelected(new Set());
      await refreshAll();
      setMessage("已删除选中记录");
    } catch (e: any) {
      setMessage(e.message);
    }
  }

  async function onClearTasks() {
    if (!window.confirm("确认清空当前全部记录吗？")) return;
    try {
      await clearAllTasks();
      setSelected(new Set());
      await refreshAll();
      setMessage("已清空全部记录");
    } catch (e: any) {
      setMessage(e.message);
    }
  }

  async function onToggleAuto() {
    const next = !autoWatch;
    try {
      await toggleAutoWatch(next);
      setAutoWatch(next);
    } catch (e: any) {
      setMessage(e.message);
    }
  }

  async function onUploadContacts(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      await uploadContacts(file);
      await refreshAll();
      setMessage("通讯录已更新");
    } catch (err: any) {
      setMessage(err.message);
    }
  }

  async function onSaveConfig() {
    try {
      await updateConfig({
        root_path: rootPath,
        corp_id: corpId || undefined,
        agent_id: agentId || undefined,
        secret: secret || undefined
      });
      await refreshAll();
      setMessage("配置已保存");
    } catch (e: any) {
      setMessage(e.message);
    }
  }

  async function onCheckIp() {
    try {
      const r = await getPublicIp();
      setPublicIp(r.ip || "-");
    } catch {
      setPublicIp("获取失败");
    }
  }

  function patchRpa<K extends keyof RpaStartPayload>(key: K, value: RpaStartPayload[K]) {
    setRpaForm((prev) => ({ ...prev, [key]: value }));
  }

  async function onStartRpa() {
    if (!window.confirm(rpaForm.paste_only ? "当前是仅粘贴模式，确认开始？" : "当前是正式发送模式，确认开始？")) return;
    setRpaLoading(true);
    try {
      const s = await startRpa(rpaForm);
      setRpaStatus(s);
      await refreshRpa();
      setRpaMessage("RPA 已启动");
    } catch (e: any) {
      setRpaMessage(e.message);
    } finally {
      setRpaLoading(false);
    }
  }

  async function onStopRpa() {
    setRpaLoading(true);
    try {
      const s = await stopRpa();
      setRpaStatus(s);
      await refreshRpa();
      setRpaMessage("RPA 已停止");
    } catch (e: any) {
      setRpaMessage(e.message);
    } finally {
      setRpaLoading(false);
    }
  }

  function applyRpaPreset(preset: "fast" | "stable" | "strong") {
    if (preset === "fast") {
      setRpaForm((p) => ({
        ...p,
        interval_sec: 2,
        timeout_sec: 12,
        max_retries: 2,
        retry_delay_sec: 1,
        stabilize_open_rounds: 1,
        stabilize_focus_rounds: 1,
        stabilize_send_rounds: 0,
        open_chat_strategy: "keyboard_first"
      }));
      return;
    }
    if (preset === "strong") {
      setRpaForm((p) => ({
        ...p,
        interval_sec: 5,
        timeout_sec: 24,
        max_retries: 3,
        retry_delay_sec: 2,
        stabilize_open_rounds: 3,
        stabilize_focus_rounds: 3,
        stabilize_send_rounds: 2,
        open_chat_strategy: "hybrid"
      }));
      return;
    }
    setRpaForm((p) => ({
      ...p,
      interval_sec: 4,
      timeout_sec: 20,
      max_retries: 3,
      retry_delay_sec: 1.5,
      stabilize_open_rounds: 2,
      stabilize_focus_rounds: 2,
      stabilize_send_rounds: 1,
      open_chat_strategy: "keyboard_first"
    }));
  }

  return (
    <div className="shell">
      <div className="hero">
        <div>
          <h1>WeCom Sender Console</h1>
          <p>同一界面管理 API 批量发送与本机 RPA 粘贴/发送，不改核心脚本</p>
        </div>
        <div className="hero-badges">
          <span className="chip">{rpaStatus.running ? "RPA 运行中" : "RPA 空闲"}</span>
          <span className="chip">任务 {status?.total ?? 0}</span>
        </div>
      </div>

      <div className="tabs">
        <button className={activeTab === "rpa" ? "tab active" : "tab"} onClick={() => setActiveTab("rpa")}>RPA 本机模式</button>
        <button className={activeTab === "api" ? "tab active" : "tab"} onClick={() => setActiveTab("api")}>API 批量模式</button>
      </div>

      {activeTab === "rpa" && (
        <div className="grid">
          <section className="card">
            <h3>运行控制</h3>
            <div className="toolbar">
              <button className="accent" onClick={onStartRpa} disabled={rpaLoading || rpaStatus.running}>开始运行</button>
              <button onClick={onStopRpa} disabled={rpaLoading || !rpaStatus.running}>停止运行</button>
              <button onClick={() => refreshRpa()}>刷新状态</button>
              <button onClick={() => navigator.clipboard.writeText(rpaCmdPreview)}>复制命令</button>
            </div>
            <div className="status-grid">
              <span>PID: {rpaStatus.pid ?? "-"}</span>
              <span>开始: {rpaStatus.started_at ?? "-"}</span>
              <span>结束: {rpaStatus.finished_at ?? "-"}</span>
              <span>退出码: {rpaStatus.return_code ?? "-"}</span>
              <span>sent: {rpaCounts.sent ?? 0}</span>
              <span>pasted: {(rpaCounts.pasted_only ?? 0) + (rpaCounts.pasted_unverified ?? 0)}</span>
              <span>pasted_unverified: {rpaCounts.pasted_unverified ?? 0}</span>
              <span>failed: {rpaCounts.failed ?? 0}</span>
            </div>
            {rpaMessage && <div className="notice">{rpaMessage}</div>}
            <pre className="cmd">{rpaCmdPreview}</pre>
          </section>

          <section className="card">
            <h3>参数配置</h3>
            <div className="preset-row">
              <span>预设</span>
              <button onClick={() => applyRpaPreset("fast")}>快</button>
              <button onClick={() => applyRpaPreset("stable")}>稳</button>
              <button onClick={() => applyRpaPreset("strong")}>抗干扰</button>
            </div>
            <div className="form-grid">
              <label>任务 CSV</label>
              <input value={rpaForm.tasks_csv} onChange={(e) => patchRpa("tasks_csv", e.target.value)} />
              <label>WeCom EXE</label>
              <input value={rpaForm.wecom_exe || ""} onChange={(e) => patchRpa("wecom_exe", e.target.value)} />
              <label>标题正则</label>
              <input value={rpaForm.main_title_re} onChange={(e) => patchRpa("main_title_re", e.target.value)} />
              <label>发送模式</label>
              <select value={rpaForm.send_mode} onChange={(e) => patchRpa("send_mode", e.target.value as RpaStartPayload["send_mode"])}>
                <option value="clipboard">clipboard</option>
                <option value="dialog">dialog</option>
                <option value="auto">auto</option>
              </select>
              <label>开会话策略</label>
              <select value={rpaForm.open_chat_strategy} onChange={(e) => patchRpa("open_chat_strategy", e.target.value as RpaStartPayload["open_chat_strategy"])}>
                <option value="keyboard_first">keyboard_first</option>
                <option value="click_first">click_first</option>
                <option value="hybrid">hybrid</option>
              </select>
              <label>任务间隔秒</label>
              <input type="number" value={rpaForm.interval_sec} onChange={(e) => patchRpa("interval_sec", Number(e.target.value))} />
              <label>超时秒</label>
              <input type="number" value={rpaForm.timeout_sec} onChange={(e) => patchRpa("timeout_sec", Number(e.target.value))} />
              <label>重试次数</label>
              <input type="number" value={rpaForm.max_retries} onChange={(e) => patchRpa("max_retries", Number(e.target.value))} />
              <label>重试间隔</label>
              <input type="number" step="0.1" value={rpaForm.retry_delay_sec} onChange={(e) => patchRpa("retry_delay_sec", Number(e.target.value))} />
              <label>稳态轮次(开会话)</label>
              <input type="number" value={rpaForm.stabilize_open_rounds} onChange={(e) => patchRpa("stabilize_open_rounds", Number(e.target.value))} />
              <label>稳态轮次(聚焦)</label>
              <input type="number" value={rpaForm.stabilize_focus_rounds} onChange={(e) => patchRpa("stabilize_focus_rounds", Number(e.target.value))} />
              <label>稳态轮次(发送前)</label>
              <input type="number" value={rpaForm.stabilize_send_rounds} onChange={(e) => patchRpa("stabilize_send_rounds", Number(e.target.value))} />
              <label>从第几条开始</label>
              <input type="number" value={rpaForm.resume_from} onChange={(e) => patchRpa("resume_from", Number(e.target.value))} />
              <label>结果文件</label>
              <input value={rpaForm.results_csv} onChange={(e) => patchRpa("results_csv", e.target.value)} />
              <label>日志文件</label>
              <input value={rpaForm.log_file} onChange={(e) => patchRpa("log_file", e.target.value)} />
            </div>
            <div className="toggle-grid">
              <label><input type="checkbox" checked={rpaForm.paste_only} onChange={(e) => patchRpa("paste_only", e.target.checked)} /> 仅粘贴</label>
              <label><input type="checkbox" checked={rpaForm.no_chat_verify} onChange={(e) => patchRpa("no_chat_verify", e.target.checked)} /> 关闭会话校验</label>
              <label><input type="checkbox" checked={rpaForm.resume_failed} onChange={(e) => patchRpa("resume_failed", e.target.checked)} /> 只续跑失败项</label>
              <label><input type="checkbox" checked={rpaForm.stop_on_fail} onChange={(e) => patchRpa("stop_on_fail", e.target.checked)} /> 失败即停</label>
              <label><input type="checkbox" checked={rpaForm.skip_missing_image} onChange={(e) => patchRpa("skip_missing_image", e.target.checked)} /> 缺图跳过</label>
              <label><input type="checkbox" checked={rpaForm.debug_chat_text} onChange={(e) => patchRpa("debug_chat_text", e.target.checked)} /> 调试会话文本</label>
              <label><input type="checkbox" checked={rpaForm.dry_run} onChange={(e) => patchRpa("dry_run", e.target.checked)} /> Dry Run</label>
            </div>
          </section>

          <section className="card span-2">
            <h3>实时日志</h3>
            <pre className="logbox">{rpaLogs.join("\n") || "暂无日志"}</pre>
          </section>
        </div>
      )}

      {activeTab === "api" && (
        <div className="grid">
          <section className="card span-2">
            <h3>API 批量模式</h3>
            <div className="toolbar">
              <input value={rootPath} onChange={(e) => setRootPath(e.target.value)} />
              <button onClick={onSaveConfig}>保存配置</button>
              <button onClick={onScan}>扫描</button>
              <button className="accent" onClick={onBatchSend}>批量发送</button>
              <button onClick={onSelectedSend} disabled={selectedCount === 0}>发送选中({selectedCount})</button>
              <button onClick={onDeleteSelected} disabled={selectedCount === 0}>删除选中({selectedCount})</button>
              <button onClick={onClearTasks} disabled={tasks.length === 0}>清空记录</button>
              <button onClick={onToggleAuto}>{autoWatch ? "关闭自动监听" : "开启自动监听"}</button>
            </div>
            <div className="status-grid">
              <span>总数: {status?.total ?? 0}</span>
              <span>待发送: {status?.pending ?? 0}</span>
              <span>排队: {status?.queued ?? 0}</span>
              <span>发送中: {status?.sending ?? 0}</span>
              <span>成功: {status?.sent ?? 0}</span>
              <span>失败: {status?.failed ?? 0}</span>
              <span>跳过: {status?.skipped ?? 0}</span>
            </div>
            {message && <div className="notice">{message}</div>}
          </section>

          <section className="card">
            <h3>企业微信配置</h3>
            <div className="form-grid">
              <label>corp_id</label>
              <input value={corpId} onChange={(e) => setCorpId(e.target.value)} placeholder={config?.corp_id || ""} />
              <label>agent_id</label>
              <input value={agentId} onChange={(e) => setAgentId(e.target.value)} placeholder={config?.agent_id || ""} />
              <label>secret</label>
              <input value={secret} onChange={(e) => setSecret(e.target.value)} placeholder={config?.secret || ""} />
            </div>
            <div className="toolbar">
              <button onClick={onCheckIp}>检测出口IP</button>
              <span className="mono">{publicIp ?? "-"}</span>
            </div>
            <div className="toolbar">
              <input type="file" accept=".xlsx" onChange={onUploadContacts} />
            </div>
          </section>

          <section className="card span-2">
            <TaskTable tasks={tasks} selected={selected} onToggle={toggleSelection} onDelete={onDeleteOne} />
          </section>
        </div>
      )}
    </div>
  );
}

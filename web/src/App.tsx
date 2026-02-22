import React, { useEffect, useMemo, useState } from "react";
import {
  getStatus,
  listTasks,
  scan,
  sendBatch,
  sendSelected,
  toggleAutoWatch,
  uploadContacts,
  getConfig,
  updateConfig,
  getPublicIp,
  type Task
} from "./api";
import TaskTable from "./components/TaskTable";

export default function App() {
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

  const selectedCount = useMemo(() => selected.size, [selected]);

  async function refreshAll() {
    const [t, s, c] = await Promise.all([listTasks(), getStatus(), getConfig()]);
    setTasks(t);
    setStatus(s);
    setConfig(c);
    setRootPath(c.root_path || "/data/inbox");
  }

  useEffect(() => {
    refreshAll().catch((e) => setMessage(e.message));
    const id = setInterval(() => {
      refreshAll().catch(() => undefined);
    }, 3000);
    return () => clearInterval(id);
  }, []);

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
      setMessage("已开始发送勾选项");
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
    } catch (e: any) {
      setPublicIp("获取失败");
    }
  }

  return (
    <div className="app">
      <header>
        <h1>企业微信批量图片发送</h1>
        <p className="muted">递归扫描文件夹 → 自动匹配妈妈 → 队列发送 → 状态监控</p>
      </header>

      <section className="panel">
        <div className="row">
          <label>根目录</label>
          <input value={rootPath} onChange={(e) => setRootPath(e.target.value)} />
          <button onClick={onSaveConfig}>保存配置</button>
          <button onClick={onScan}>扫描</button>
          <button className="primary" onClick={onBatchSend}>批量发送</button>
          <button onClick={onSelectedSend} disabled={selectedCount === 0}>
            发送勾选({selectedCount})
          </button>
          <button onClick={onToggleAuto}>
            {autoWatch ? "关闭自动监控" : "开启自动监控"}
          </button>
        </div>
        <div className="row">
          <label>corp_id</label>
          <input value={corpId} onChange={(e) => setCorpId(e.target.value)} placeholder={config?.corp_id || ""} />
          <label>agent_id</label>
          <input value={agentId} onChange={(e) => setAgentId(e.target.value)} placeholder={config?.agent_id || ""} />
          <label>secret</label>
          <input value={secret} onChange={(e) => setSecret(e.target.value)} placeholder={config?.secret || ""} />
          <button onClick={onCheckIp}>检测出口 IP</button>
          <span className="mono">{publicIp ?? "-"}</span>
        </div>
        <div className="row">
          <label>通讯录 Excel</label>
          <input type="file" accept=".xlsx" onChange={onUploadContacts} />
        </div>
        <div className="row">
          <label>当前配置</label>
          <span className="mono">corp_id: {config?.corp_id || "-"}</span>
          <span className="mono">agent_id: {config?.agent_id || "-"}</span>
          <span className="mono">secret: {config?.secret ? "已保存" : "-"}</span>
        </div>
      </section>

      <section className="panel">
        <div className="status">
          <span>总数: {status?.total ?? 0}</span>
          <span>待发送: {status?.pending ?? 0}</span>
          <span>排队: {status?.queued ?? 0}</span>
          <span>发送中: {status?.sending ?? 0}</span>
          <span>成功: {status?.sent ?? 0}</span>
          <span>失败: {status?.failed ?? 0}</span>
          <span>跳过: {status?.skipped ?? 0}</span>
        </div>
        {message && <div className="toast">{message}</div>}
      </section>

      <TaskTable tasks={tasks} selected={selected} onToggle={toggleSelection} />
    </div>
  );
}

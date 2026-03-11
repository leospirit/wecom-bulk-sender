(function () {
  const e = React.createElement;
  const useEffect = React.useEffect;
  const useMemo = React.useMemo;
  const useState = React.useState;

  const DEFAULT_FORM = {
    tasks_csv: "tools/rpa_tasks.real.csv",
    send_mode: "clipboard",
    main_title_re: ".*(WeCom|WXWork|企业微信).*",
    interval_sec: 4,
    timeout_sec: 20,
    max_retries: 3,
    retry_delay_sec: 1.5,
    stabilize_open_rounds: 2,
    stabilize_focus_rounds: 2,
    stabilize_send_rounds: 1,
    open_chat_strategy: "keyboard_first",
    resume_from: 1,
    paste_only: true,
    no_chat_verify: true,
    resume_failed: false,
    stop_on_fail: true,
    skip_missing_image: true,
    debug_chat_text: false,
    dry_run: false,
    results_csv: "run-logs/rpa-results.csv",
    log_file: "run-logs/rpa-sender.log"
  };

  function fetchJson(url, options) {
    return fetch(url, options).then(async (r) => {
      if (!r.ok) {
        let msg = "request failed";
        try {
          const j = await r.json();
          msg = j.detail || msg;
        } catch (_e) {}
        throw new Error(msg);
      }
      return r.json();
    });
  }

  function App() {
    const [form, setForm] = useState(DEFAULT_FORM);
    const [status, setStatus] = useState({ running: false, result_counts: {} });
    const [logs, setLogs] = useState([]);
    const [msg, setMsg] = useState("就绪");
    const [busy, setBusy] = useState(false);

    const counts = status.result_counts || {};
    const pasted = (counts.pasted_only || 0) + (counts.pasted_unverified || 0);

    const cmdPreview = useMemo(() => {
      const cmd = [
        "python tools/wecom_rpa_sender.py",
        "--tasks-csv " + form.tasks_csv,
        "--send-mode " + form.send_mode,
        "--interval-sec " + form.interval_sec,
        "--timeout-sec " + form.timeout_sec,
        "--max-retries " + form.max_retries,
        "--retry-delay-sec " + form.retry_delay_sec
      ];
      if (form.paste_only) cmd.push("--paste-only");
      if (form.no_chat_verify) cmd.push("--no-chat-verify");
      if (form.resume_failed) cmd.push("--resume-failed");
      if (form.stop_on_fail) cmd.push("--stop-on-fail");
      return cmd.join(" ");
    }, [form]);

    function patch(key, value) {
      setForm((p) => Object.assign({}, p, { [key]: value }));
    }

    async function refresh() {
      const [s, l] = await Promise.all([
        fetchJson("/api/rpa/status"),
        fetchJson("/api/rpa/log-tail?lines=180")
      ]);
      setStatus(s);
      setLogs(l.lines || []);
    }

    async function startRun() {
      if (!confirm(form.paste_only ? "仅粘贴模式，确认开始？" : "正式发送模式，确认开始？")) return;
      setBusy(true);
      try {
        await fetchJson("/api/rpa/start", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(form)
        });
        setMsg("已启动");
        await refresh();
      } catch (err) {
        setMsg(err.message || String(err));
      } finally {
        setBusy(false);
      }
    }

    async function stopRun() {
      setBusy(true);
      try {
        await fetchJson("/api/rpa/stop", { method: "POST" });
        setMsg("已停止");
        await refresh();
      } catch (err) {
        setMsg(err.message || String(err));
      } finally {
        setBusy(false);
      }
    }

    useEffect(() => {
      refresh().catch(() => undefined);
      const id = setInterval(() => refresh().catch(() => undefined), 2500);
      return () => clearInterval(id);
    }, []);

    return e(
      "div",
      { className: "wrap" },
      e("div", { className: "hero" }, e("h1", null, "RPA Lite Console"), e("p", null, "不依赖 Node/Vite，避免 EPERM；只启动本机 FastAPI，不影响 Docker 其他程序")),
      e(
        "div",
        { className: "card" },
        e(
          "div",
          { className: "toolbar" },
          e("button", { className: "primary", onClick: startRun, disabled: busy || status.running }, "开始运行"),
          e("button", { onClick: stopRun, disabled: busy || !status.running }, "停止运行"),
          e("button", { onClick: () => refresh().catch(() => undefined) }, "刷新"),
          e("span", null, status.running ? "运行中" : "空闲")
        ),
        e(
          "div",
          { className: "grid" },
          e("label", null, "任务 CSV"),
          e("input", { value: form.tasks_csv, onChange: (ev) => patch("tasks_csv", ev.target.value) }),
          e("label", null, "发送模式"),
          e(
            "select",
            { value: form.send_mode, onChange: (ev) => patch("send_mode", ev.target.value) },
            e("option", { value: "clipboard" }, "clipboard"),
            e("option", { value: "dialog" }, "dialog"),
            e("option", { value: "auto" }, "auto")
          ),
          e("label", null, "间隔秒"),
          e("input", { type: "number", value: form.interval_sec, onChange: (ev) => patch("interval_sec", Number(ev.target.value)) }),
          e("label", null, "超时秒"),
          e("input", { type: "number", value: form.timeout_sec, onChange: (ev) => patch("timeout_sec", Number(ev.target.value)) }),
          e("label", null, "最大重试"),
          e("input", { type: "number", value: form.max_retries, onChange: (ev) => patch("max_retries", Number(ev.target.value)) }),
          e("label", null, "仅粘贴"),
          e("input", { type: "checkbox", checked: !!form.paste_only, onChange: (ev) => patch("paste_only", ev.target.checked) })
        ),
        e(
          "div",
          { className: "status" },
          e("div", null, "PID: ", status.pid || "-"),
          e("div", null, "sent: ", counts.sent || 0),
          e("div", null, "pasted: ", pasted),
          e("div", null, "failed: ", counts.failed || 0)
        ),
        e("div", { className: "hint" }, msg),
        e("div", { className: "hint" }, "命令预览: ", cmdPreview)
      ),
      e("div", { className: "card" }, e("h3", null, "日志"), e("div", { className: "log" }, (logs || []).join("\n") || "暂无日志"))
    );
  }

  ReactDOM.createRoot(document.getElementById("root")).render(e(App));
})();

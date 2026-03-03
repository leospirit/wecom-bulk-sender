export type Task = {
  id: number;
  file_path: string;
  student_name: string;
  parent_name: string;
  user_id: string | null;
  status: string;
  error: string | null;
  created_at: string;
  updated_at: string;
};

export type StatusCounts = {
  total: number;
  pending: number;
  queued: number;
  sending: number;
  sent: number;
  failed: number;
  skipped: number;
};

export async function getStatus(): Promise<StatusCounts> {
  const r = await fetch("/api/status");
  if (!r.ok) throw new Error("status failed");
  return r.json();
}

export async function listTasks(): Promise<Task[]> {
  const r = await fetch("/api/tasks");
  if (!r.ok) throw new Error("list tasks failed");
  return r.json();
}

export async function scan(rootPath: string): Promise<void> {
  const r = await fetch("/api/scan", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ rootPath })
  });
  if (!r.ok) throw new Error("scan failed");
}

export async function sendBatch(): Promise<void> {
  const r = await fetch("/api/send/batch", { method: "POST" });
  if (!r.ok) throw new Error("send batch failed");
}

export async function sendSelected(taskIds: number[]): Promise<void> {
  const r = await fetch("/api/send/selected", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ taskIds })
  });
  if (!r.ok) throw new Error("send selected failed");
}

export async function deleteTask(taskId: number): Promise<void> {
  const r = await fetch(`/api/tasks/${taskId}`, { method: "DELETE" });
  if (!r.ok) throw new Error("delete task failed");
}

export async function deleteSelectedTasks(taskIds: number[]): Promise<void> {
  const r = await fetch("/api/tasks/delete-selected", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ taskIds })
  });
  if (!r.ok) throw new Error("delete selected tasks failed");
}

export async function clearAllTasks(): Promise<void> {
  const r = await fetch("/api/tasks/clear", { method: "POST" });
  if (!r.ok) throw new Error("clear tasks failed");
}

export async function toggleAutoWatch(enabled: boolean): Promise<void> {
  const r = await fetch("/api/auto-watch", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ enabled })
  });
  if (!r.ok) throw new Error("toggle auto watch failed");
}

export async function uploadContacts(file: File): Promise<void> {
  const data = new FormData();
  data.append("file", file);
  const r = await fetch("/api/contacts/upload", { method: "POST", body: data });
  if (!r.ok) throw new Error("upload contacts failed");
}

export type AppConfig = {
  corp_id: string;
  agent_id: string;
  secret: string;
  root_path: string;
  rate_limit_per_sec: number;
  max_concurrency: number;
};

export async function getConfig(): Promise<AppConfig> {
  const r = await fetch("/api/config");
  if (!r.ok) throw new Error("get config failed");
  return r.json();
}

export async function updateConfig(cfg: Partial<AppConfig>): Promise<void> {
  const r = await fetch("/api/config", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(cfg)
  });
  if (!r.ok) throw new Error("update config failed");
}

export async function getPublicIp(): Promise<{ ip: string }> {
  const r = await fetch("/api/ip");
  if (!r.ok) throw new Error("get ip failed");
  return r.json();
}

export type RpaOpenChatStrategy = "keyboard_first" | "click_first" | "hybrid";
export type RpaSendMode = "clipboard" | "dialog" | "auto";

export type RpaStartPayload = {
  tasks_csv: string;
  wecom_exe?: string;
  main_title_re: string;
  send_mode: RpaSendMode;
  dry_run: boolean;
  paste_only: boolean;
  no_chat_verify: boolean;
  interval_sec: number;
  timeout_sec: number;
  max_retries: number;
  retry_delay_sec: number;
  stabilize_open_rounds: number;
  stabilize_focus_rounds: number;
  stabilize_send_rounds: number;
  open_chat_strategy: RpaOpenChatStrategy;
  resume_from: number;
  resume_failed: boolean;
  stop_on_fail: boolean;
  skip_missing_image: boolean;
  debug_chat_text: boolean;
  results_csv: string;
  log_file: string;
};

export type RpaStatus = {
  running: boolean;
  pid: number | null;
  started_at: string | null;
  finished_at: string | null;
  return_code: number | null;
  command: string[];
  log_file: string;
  results_csv: string;
  result_counts: Record<string, number>;
};

export async function getRpaStatus(): Promise<RpaStatus> {
  const r = await fetch("/api/rpa/status");
  if (!r.ok) throw new Error("get rpa status failed");
  return r.json();
}

export async function startRpa(payload: RpaStartPayload): Promise<RpaStatus> {
  const r = await fetch("/api/rpa/start", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!r.ok) {
    let msg = "start rpa failed";
    try {
      const body = await r.json();
      msg = body?.detail || msg;
    } catch {
      // ignore parse error
    }
    throw new Error(msg);
  }
  return r.json();
}

export async function stopRpa(): Promise<RpaStatus> {
  const r = await fetch("/api/rpa/stop", { method: "POST" });
  if (!r.ok) throw new Error("stop rpa failed");
  return r.json();
}

export async function getRpaLogTail(lines = 180): Promise<{ path: string; lines: string[] }> {
  const r = await fetch(`/api/rpa/log-tail?lines=${lines}`);
  if (!r.ok) throw new Error("get rpa log tail failed");
  return r.json();
}

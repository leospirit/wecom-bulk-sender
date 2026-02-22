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

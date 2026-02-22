import React from "react";
import type { Task } from "../api";

type Props = {
  tasks: Task[];
  selected: Set<number>;
  onToggle: (id: number) => void;
};

export default function TaskTable({ tasks, selected, onToggle }: Props) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th></th>
            <th>文件</th>
            <th>学生</th>
            <th>家长</th>
            <th>UserId</th>
            <th>状态</th>
            <th>原因</th>
          </tr>
        </thead>
        <tbody>
          {tasks.map((t) => (
            <tr key={t.id}>
              <td>
                <input
                  type="checkbox"
                  checked={selected.has(t.id)}
                  onChange={() => onToggle(t.id)}
                />
              </td>
              <td className="mono">{t.file_path}</td>
              <td>{t.student_name}</td>
              <td>{t.parent_name}</td>
              <td className="mono">{t.user_id ?? "-"}</td>
              <td>
                <span className={`badge ${t.status}`}>{t.status}</span>
              </td>
              <td className="muted">{t.error ?? ""}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

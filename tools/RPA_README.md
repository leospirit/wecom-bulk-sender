# WeCom RPA Sender

## 1) Install

```powershell
cd C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main
python -m pip install -r tools\rpa-requirements.txt
```

## 2) Prepare tasks CSV

Copy `tools\rpa_tasks.sample.csv` and fill rows.
Required columns:

- `parent_name`
- `image_path`

Optional:

- `student_name`
- `search_keyword` (用于搜索会话，默认等于 `parent_name`)
- `confirm_keyword` (用于发前校验当前会话，默认等于 `parent_name`)

## 3) Dry run first

```powershell
python tools\wecom_rpa_sender.py --tasks-csv tools\rpa_tasks.sample.csv --dry-run
```

## 4) Real run

```powershell
python tools\wecom_rpa_sender.py --tasks-csv tools\rpa_tasks.sample.csv
```

## 5) UI run (recommended)

Fastest way on a new Windows PC:

- double click `..\one-click-rpa.bat`

This will auto-create venv, install dependencies, and launch RPA UI.

Double click:

- `tools\run-rpa-ui.bat`

Or run in terminal:

```powershell
python tools\wecom_rpa_gui.py
```

UI supports:

- choose tasks csv
- one-click mode presets (Test / Send / Resume Failed)
- paste-only / real-send switch
- retries / interval / timeout
- resume-from / resume-failed
- live runtime logs and stop button
- Chinese 3-step flow UI: choose CSV -> choose mode -> run
- modern button set (primary/secondary/danger) + mode segmented buttons
- light/dark theme toggle
- run summary chips + completion popup (total/success/failed/skipped/time)

## 6) React UI run (new)

Double click:

- `tools\run-rpa-react-ui.bat`

Then open:

- `http://localhost:5173`

React UI includes:

- API mode + RPA mode tabs
- start/stop RPA from browser
- anti-disturb tuning (open/focus/send rounds + open-chat strategy)
- live log tail + command preview
- result counters (including `pasted_unverified`)

## Useful args

- `--wecom-exe "C:\Program Files (x86)\WXWork\WXWork.exe"`
- `--main-title-re "企业微信"`
- `--interval-sec 1.5`
- `--timeout-sec 12`
- `--results-csv run-logs\rpa-results.csv`
- `--log-file run-logs\rpa-sender.log`
- `--debug-chat-text`
- `--no-chat-verify`
- `--paste-only` (只粘贴到输入框，不发送)
- `--max-retries 2` (每条失败自动重试次数)
- `--retry-delay-sec 1.0` (重试间隔)
- `--stop-on-fail` (一条失败就停止整批)
- `--skip-missing-image` (图片缺失时跳过而非失败)
- `--resume-from 5` (从第5条任务开始，1-based)
- `--resume-failed` (只跑上次结果里非ok任务/未跑任务)
- `--resume-results-csv run-logs\rpa-results.csv` (续跑参考结果文件)
- `--stabilize-open-rounds 2` (开会话前稳态ESC轮次)
- `--stabilize-focus-rounds 2` (聚焦输入框前稳态ESC轮次)
- `--stabilize-send-rounds 1` (粘贴/发送前稳态ESC轮次)
- `--open-chat-strategy keyboard_first` (会话打开策略: keyboard_first/click_first/hybrid)

## Notes

- Keep WeCom in foreground during run.
- Do not touch keyboard/mouse while running.
- If open-file dialog is not found, adjust shortcut/selector in `send_image_via_file_dialog`.

## 7) 300+ Batch SOP (Stable)

### A. Pre-run checklist

1. WeCom desktop is logged in and visible.
2. `tools\rpa_tasks.real.csv` has been spot-checked (name/path mapping).
3. Image files exist and paths are valid.
4. Disable desktop distractions as much as possible:
   - mute notifications / popups
   - keep WeCom on top during run
5. Use `tools\run-rpa-ui.bat` (Tkinter runner as primary stable entry).

### B. Recommended stable parameters

- `send_mode = clipboard`
- `interval_sec = 4`
- `timeout_sec = 20`
- `max_retries = 3`
- `retry_delay_sec = 1.5`
- `stop_on_fail = on`
- `skip_missing_image = on`
- `paste_only = on` for rehearsal, then `off` for real send
- anti-disturb:
  - `stabilize_open_rounds = 2`
  - `stabilize_focus_rounds = 2`
  - `stabilize_send_rounds = 1`
  - `open_chat_strategy = keyboard_first`

### C. Run process

1. Rehearsal: run paste-only first.
2. Spot-check 5-10 chats for correct chat target and image.
3. Switch to real send mode.
4. Run in chunks of `50-80` tasks (not all 300+ at once).

### D. Failure handling

1. If one task fails, stop immediately (`stop_on_fail`).
2. Fix root cause (focus/window/network/missing image).
3. Use `resume_failed` to continue only unfinished/failed items.
4. Repeat until `failed=0`.

### E. Completion checklist

1. `run-logs\rpa-results.csv` has no `failed`.
2. Randomly sample at least 10 recipients from different chunks.
3. Archive logs:
   - `run-logs\rpa-sender.log`
   - `run-logs\rpa-results.csv`

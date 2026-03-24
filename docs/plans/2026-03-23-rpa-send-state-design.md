# CSV Send State Design

This change adds a minimal send-state ledger so WeCom refresh/build logic can answer "what still needs to be pushed" instead of only "what changed in the last refresh".

## Problem

`rpa_tasks.delta.csv` only captures rows whose `report_submission_id` changed compared with the previous refresh snapshot. It does not know whether an older delta batch was actually pasted/sent. In real use, reports are often produced in multiple batches. After the last refresh, delta may contain only the newest batch, while earlier unsent rows disappear from the active send set.

## Goal

Introduce a CSV state ledger that tracks handled task keys and generate a `pending` task CSV containing current rows that have not yet been handled.

## Scope

- Keep existing full snapshot: `tools/rpa_tasks.with-message.csv`
- Keep existing change snapshot: `tools/rpa_tasks.delta.csv`
- Add new pending snapshot: `tools/rpa_tasks.pending.csv`
- Add new state ledger: `run-logs/rpa_send_state.csv`
- Default RPA execution should target `pending.csv`

## State Model

Stable task key:
- `class_name`
- `student_name`
- `report_submission_id`

Stored state fields:
- `class_name`
- `student_name`
- `report_submission_id`
- `parent_name`
- `image_path`
- `status`
- `text_status`
- `updated_at`

Handled statuses for pending filtering:
- `ok`
- `sent`
- `pasted_only`
- `pasted_unverified`

Statuses such as `failed`, `dry_run`, and `skipped_missing_image` stay pending on future refreshes.

## Data Flow

1. Refresh script builds current full rows as today.
2. Refresh script reads `rpa_send_state.csv` if present.
3. Refresh script writes:
   - full snapshot
   - delta snapshot
   - pending snapshot = full rows minus handled state keys
4. Sender writes/updates send-state rows after each task run.
5. API/web/local GUI default `tasks_csv` switches from delta to pending.

## Safety

- No database introduced.
- No deletion of historical full or delta snapshots.
- New report version always becomes pending again because `report_submission_id` changes.
- Existing manual override of `tasks_csv` remains available.

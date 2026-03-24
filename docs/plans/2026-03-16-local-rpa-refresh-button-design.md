# Local RPA Refresh Button Design

**Date:** 2026-03-16

## Goal

Add a `刷新发送任务` button to the local Tkinter GUI so the desktop workflow matches the web workflow:

1. User manually drops new reports into the speech scoring `data/out` directory.
2. User clicks one refresh button in the local GUI.
3. The GUI rebuilds `tools/rpa_tasks.with-message.csv` from reports + contacts.
4. The refreshed CSV becomes the active task file and can be sent immediately.

## Constraints

- Reuse the existing report-driven rebuild logic.
- Do not duplicate matching or recommendation logic in the GUI.
- Keep the current matching rules unchanged:
  - class first, student name fallback
  - mother first, father fallback
  - latest report per student only
- Use the existing contacts file:
  - `data/contacts.xlsx`
- Use the existing speech scoring API:
  - `http://127.0.0.1:8000`

## Recommended Approach

Add a refresh button to the Step 1 task-file card in `tools/wecom_rpa_gui.py` and have it invoke the existing script:

`tools/build_rpa_tasks_from_score.py --report-driven --contacts-xlsx data/contacts.xlsx --output-csv tools/rpa_tasks.with-message.csv --score-api-base http://127.0.0.1:8000`

This is the most stable option because it reuses the same pipeline already verified in the web version.

## UI Behavior

- Change the local GUI default task CSV from `tools/rpa_tasks.real.csv` to `tools/rpa_tasks.with-message.csv`.
- Add a new button near `选择` and `检查任务`:
  - `↻ 刷新发送任务`
- When clicked:
  - run the report-driven refresh command in a background thread
  - write progress and errors to the log pane
  - on success:
    - set `csv_path` to `tools/rpa_tasks.with-message.csv`
    - run `检查任务`
    - update the status text with the rebuild summary
- While refresh is running:
  - disable the refresh button and run button
  - keep stop behavior unchanged for actual send tasks

## Error Handling

- Missing `data/contacts.xlsx`: show a message box and log the failure
- Missing rebuild script: show a message box and log the failure
- Speech scoring API unavailable: log stderr/stdout and show concise status failure
- Rebuild succeeds with zero rows: still treat as success, but show summary clearly

## Testing Strategy

Use TDD and extract the refresh command + refresh summary parsing into testable helpers inside the GUI module.

Test coverage:

1. Local GUI default CSV points to `tools/rpa_tasks.with-message.csv`
2. Refresh command points to:
   - `build_rpa_tasks_from_score.py`
   - `data/contacts.xlsx`
   - `tools/rpa_tasks.with-message.csv`
   - `http://127.0.0.1:8000`
3. Refresh summary parsing handles:
   - `total`
   - `enriched`
   - `no_match_or_empty`
   - `errors`
4. GUI success path updates `csv_path` to the rebuilt CSV

## Files Expected To Change

- Modify: `tools/wecom_rpa_gui.py`
- Create or modify tests under: `tools/tests/`


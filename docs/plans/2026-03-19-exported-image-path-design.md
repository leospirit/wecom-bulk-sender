# Exported Image Path Matching Design

**Date:** 2026-03-19

## Goal

Make `rpa_tasks.with-message.csv` prefer real exported report PNGs under the score-reading output directory instead of assuming images are stored as `<submission_id>.png` next to report HTML.

## Problem

The current report-driven refresh script derives `image_path` from the report HTML path or report id. That only works when a PNG named after `submission_id` exists beside the report HTML. In practice, the user exports report PNGs manually into `D:\score_reading_fresh\data\out` using human-readable filenames such as `李宛迅_六年级8班_一单元_report.png`. As a result, generated tasks point at missing files and WeCom skips every send.

## Constraints

- Keep the change isolated to the WeCom sender side.
- Do not require changes to report scoring logic.
- Preserve existing fallback behavior when no exported PNG is found.
- Keep existing contact matching rules intact:
  - class first, name fallback
  - mother first, father fallback
  - latest report per student

## Recommended Approach

Scan the score-reading export directory for real PNGs during `--report-driven` refresh, build an index keyed by normalized student/class identity, and use that index to populate `image_path`. Fall back to the old `<submission_id>.png` derivation only when no exported image can be matched.

## Matching Rules

### Exported image discovery

- Scan configurable export image root, defaulting to `D:\score_reading_fresh\data\out` for the current machine.
- Only consider files matching `*_report.png`.
- Ignore non-report PNGs.

### Filename parsing

Parse the filename stem before `_report` or `+report`.

Examples:
- `李宛迅_六年级8班_一单元_report.png`
- `李宛迅+六年级8班+一单元_report.png`

Extract:
- student name: first segment
- class name: second segment if it looks like a class marker such as `*班`

### Matching priority

For each report/task:
1. try exact `class + student` match against discovered exported PNGs
2. if unavailable, try `student`-only match
3. if multiple candidate PNGs remain, prefer newest file modification time
4. if still no match, fall back to existing `_html_to_image_path(...)`

## Configuration

Add optional CLI parameter to refresh script:
- `--export-images-dir`

Default behavior:
- when omitted, use the score-reading output directory already used in this environment
- keep the old derived path fallback so the script remains usable elsewhere

## Expected Behavior

After the change:
1. user exports report PNGs into the score-reading output directory
2. user clicks refresh in the WeCom tool
3. generated `rpa_tasks.with-message.csv` points `image_path` at real exported PNGs
4. WeCom sender no longer logs `Skip missing image` for those reports

## Files Likely Affected

- `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py`
- `C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\tests\...`
- possibly GUI/API refresh entrypoints if they need to pass the new export image directory

## Testing Strategy

1. Add unit tests for filename parsing and exported image indexing.
2. Add a refresh test proving real exported PNGs override derived `<submission_id>.png` paths.
3. Add fallback test proving old behavior still works when no exported PNG is found.
4. Re-run existing refresh-related tests.

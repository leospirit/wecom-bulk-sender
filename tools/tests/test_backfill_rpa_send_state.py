from pathlib import Path
import importlib.util


MODULE_PATH = Path(
    r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\backfill_rpa_send_state.py"
)
spec = importlib.util.spec_from_file_location("backfill_rpa_send_state", MODULE_PATH)
mod = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(mod)


def test_build_backfill_state_rows_matches_exact_result_to_current_task():
    result_rows = [
        {
            "parent_name": "柯其昕爸爸",
            "student_name": "柯其昕",
            "image_path": r"D:\old\柯其昕_六年级4班_一单元_report.png",
            "message_text": "家长您好！重点词：Guinness",
            "status": "pasted_unverified",
            "text_status": "Paste not UIA-confirmed",
            "timestamp": "2026-03-23T11:56:45",
        }
    ]
    task_rows = [
        {
            "parent_name": "柯其昕爸爸",
            "student_name": "柯其昕",
            "class_name": "六年级4班",
            "report_submission_id": "web_20260323032324_907ca13a",
            "image_path": r"D:\score_reading_fresh\data\out\六年级\六年级4班\柯其昕_六年级4班_一单元_report.png",
            "message_text": "家长您好！重点词：Guinness",
        }
    ]

    rows, summary = mod.build_backfill_state_rows(result_rows, task_rows)

    assert rows == [
        {
            "parent_name": "柯其昕爸爸",
            "student_name": "柯其昕",
            "class_name": "六年级4班",
            "report_submission_id": "web_20260323032324_907ca13a",
            "image_path": r"D:\score_reading_fresh\data\out\六年级\六年级4班\柯其昕_六年级4班_一单元_report.png",
            "status": "pasted_unverified",
            "text_status": "Paste not UIA-confirmed",
            "updated_at": "2026-03-23T11:56:45",
        }
    ]
    assert summary == {
        "handled_results": 1,
        "matched": 1,
        "unmatched": 0,
        "ambiguous": 0,
        "ignored_status": 0,
    }


def test_build_backfill_state_rows_skips_unhandled_and_unmatched_rows():
    result_rows = [
        {
            "parent_name": "王小成妈妈",
            "student_name": "王小成",
            "image_path": r"D:\old\web_20260322122414_7eda2ecc.png",
            "message_text": "msg-1",
            "status": "skipped_missing_image",
            "text_status": "Image not found",
            "timestamp": "2026-03-23T11:57:03",
        },
        {
            "parent_name": "刘子芮一妈妈",
            "student_name": "刘子芮一",
            "image_path": r"D:\old\刘子芮一_六年级4班_一单元_report.png",
            "message_text": "msg-2",
            "status": "pasted_only",
            "text_status": "ok",
            "timestamp": "2026-03-23T11:57:06",
        },
    ]
    task_rows = []

    rows, summary = mod.build_backfill_state_rows(result_rows, task_rows)

    assert rows == []
    assert summary == {
        "handled_results": 1,
        "matched": 0,
        "unmatched": 1,
        "ambiguous": 0,
        "ignored_status": 1,
    }


def test_build_backfill_state_rows_falls_back_to_unique_parent_student_image_match():
    result_rows = [
        {
            "parent_name": "柯其昕爸爸",
            "student_name": "柯其昕",
            "image_path": r"D:\old\柯其昕_六年级4班_一单元_report.png",
            "message_text": "旧签名链接",
            "status": "pasted_unverified",
            "text_status": "ok",
            "timestamp": "2026-03-23T11:56:45",
        }
    ]
    task_rows = [
        {
            "parent_name": "柯其昕爸爸",
            "student_name": "柯其昕",
            "class_name": "六年级4班",
            "report_submission_id": "web_20260323032324_907ca13a",
            "image_path": r"D:\score_reading_fresh\data\out\六年级\六年级4班\柯其昕_六年级4班_一单元_report.png",
            "message_text": "新签名链接",
        }
    ]

    rows, summary = mod.build_backfill_state_rows(result_rows, task_rows)

    assert rows[0]["report_submission_id"] == "web_20260323032324_907ca13a"
    assert summary == {
        "handled_results": 1,
        "matched": 1,
        "unmatched": 0,
        "ambiguous": 0,
        "ignored_status": 0,
    }


def test_merge_state_rows_replaces_existing_key():
    previous_rows = [
        {
            "parent_name": "A",
            "student_name": "学生A",
            "class_name": "六年级1班",
            "report_submission_id": "web_old",
            "image_path": "old.png",
            "status": "sent",
            "text_status": "ok",
            "updated_at": "2026-03-20T10:00:00",
        }
    ]
    new_rows = [
        {
            "parent_name": "A",
            "student_name": "学生A",
            "class_name": "六年级1班",
            "report_submission_id": "web_old",
            "image_path": "new.png",
            "status": "pasted_unverified",
            "text_status": "new",
            "updated_at": "2026-03-23T11:00:00",
        }
    ]

    assert mod.merge_state_rows(previous_rows, new_rows) == new_rows

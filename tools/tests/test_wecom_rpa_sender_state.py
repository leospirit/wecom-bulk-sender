from pathlib import Path
import importlib.util
import sys


MODULE_PATH = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_sender.py")
spec = importlib.util.spec_from_file_location("wecom_rpa_sender", MODULE_PATH)
mod = importlib.util.module_from_spec(spec)
assert spec and spec.loader
sys.modules[spec.name] = mod
spec.loader.exec_module(mod)


def test_state_rows_only_include_handled_statuses():
    rows = [
        {
            "parent_name": "陈家瑜妈妈",
            "student_name": "陈家瑜",
            "class_name": "六年级2班",
            "report_submission_id": "rep-1",
            "image_path": r"D:\x\rep-1.png",
            "status": "pasted_only",
            "text_status": "pasted_only",
        },
        {
            "parent_name": "王小明妈妈",
            "student_name": "王小明",
            "class_name": "六年级2班",
            "report_submission_id": "rep-2",
            "image_path": r"D:\x\rep-2.png",
            "status": "failed",
            "text_status": "pending",
        },
    ]
    state_rows = mod.build_send_state_rows(rows, timestamp="2026-03-23T00:00:00")
    assert state_rows == [
        {
            "parent_name": "陈家瑜妈妈",
            "student_name": "陈家瑜",
            "class_name": "六年级2班",
            "report_submission_id": "rep-1",
            "image_path": r"D:\x\rep-1.png",
            "status": "pasted_only",
            "text_status": "pasted_only",
            "updated_at": "2026-03-23T00:00:00",
        }
    ]


def test_merge_state_rows_replaces_existing_key():
    previous_rows = [
        {
            "parent_name": "陈家瑜妈妈",
            "student_name": "陈家瑜",
            "class_name": "六年级2班",
            "report_submission_id": "rep-1",
            "image_path": r"D:\x\rep-1.png",
            "status": "pasted_only",
            "text_status": "pasted_only",
            "updated_at": "old",
        }
    ]
    new_rows = [
        {
            "parent_name": "陈家瑜妈妈",
            "student_name": "陈家瑜",
            "class_name": "六年级2班",
            "report_submission_id": "rep-1",
            "image_path": r"D:\x\rep-1.png",
            "status": "sent",
            "text_status": "sent",
            "updated_at": "new",
        }
    ]
    assert mod.merge_state_rows(previous_rows, new_rows) == new_rows

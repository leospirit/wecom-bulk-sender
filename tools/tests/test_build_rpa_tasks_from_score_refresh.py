from pathlib import Path
from tempfile import TemporaryDirectory
import importlib.util

MODULE_PATH = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\build_rpa_tasks_from_score.py")
spec = importlib.util.spec_from_file_location("build_rpa_tasks_from_score", MODULE_PATH)
mod = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(mod)


def test_choose_parent_prefers_mother_over_father():
    rows = [
        {"姓名": "陈家瑜爸爸", "账号": "dad-1", "部门": "A/六年级2班/六年级2班家长"},
        {"姓名": "陈家瑜妈妈", "账号": "mom-1", "部门": "A/六年级2班/六年级2班家长"},
    ]
    idx = mod._build_contact_indexes(rows)
    chosen = idx["by_student"]["陈家瑜"]
    assert chosen["姓名"] == "陈家瑜妈妈"
    assert chosen["账号"] == "mom-1"


def test_contact_index_supports_slash_separated_student_aliases():
    rows = [
        {"姓名": "李钰昕/李钰涵妈妈", "账号": "mom-1", "部门": "A/六年级8班/六年级8班家长"},
    ]
    idx = mod._build_contact_indexes(rows)
    assert idx["by_student"]["李钰昕"]["账号"] == "mom-1"
    assert idx["by_student"]["李钰涵"]["账号"] == "mom-1"
    assert idx["by_class_student"][("六年级8班", "李钰昕")]["账号"] == "mom-1"


def test_class_and_student_match_beats_plain_name_match():
    rows = [
        {"姓名": "王小明妈妈", "账号": "mom-a", "部门": "A/六年级1班/六年级1班家长"},
        {"姓名": "王小明妈妈", "账号": "mom-b", "部门": "A/六年级2班/六年级2班家长"},
    ]
    idx = mod._build_contact_indexes(rows)
    chosen = mod._match_contact_for_report({"student_name": "王小明", "class_name": "六年级2班"}, idx)
    assert chosen is not None
    assert chosen["账号"] == "mom-b"


def test_class_and_student_match_supports_slash_separated_parent_names():
    rows = [
        {"姓名": "张明琪/张明瑄爸爸", "账号": "dad-1", "部门": "A/六年级8班/六年级8班家长"},
        {"姓名": "张明琪/张明瑄妈妈", "账号": "mom-1", "部门": "A/六年级8班/六年级8班家长"},
    ]
    idx = mod._build_contact_indexes(rows)
    chosen = mod._match_contact_for_report({"student_name": "张明琪", "class_name": "六年级8班"}, idx)
    assert chosen is not None
    assert chosen["账号"] == "mom-1"


def test_extract_class_name_prefers_first_department_group():
    dept = "A/六年级8班/六年级8班家长;B/一年级6班/一年级6班家长"
    assert mod._extract_class_name(dept) == "六年级8班"


def test_latest_report_selected_per_student_identity():
    reports = [
        {"id": "old", "student_name": "王小明", "class_name": "六年级2班", "timestamp": 100},
        {"id": "new", "student_name": "王小明", "class_name": "六年级2班", "timestamp": 200},
        {"id": "other", "student_name": "李雷", "class_name": "六年级2班", "timestamp": 150},
    ]
    latest = mod._latest_reports_by_student(reports)
    assert [r["id"] for r in latest] == ["new", "other"]


def test_build_rows_from_reports_skips_unmatched_and_builds_expected_fields():
    reports = [
        {
            "id": "rep-1",
            "student_name": "陈家瑜",
            "class_name": "六年级2班",
            "html_path": r"D:\score_reading_fresh\data\out\02\陈家瑜\rep-1\rep-1.html",
            "timestamp": 200,
        },
        {
            "id": "rep-2",
            "student_name": "不存在",
            "class_name": "六年级2班",
            "html_path": r"D:\score_reading_fresh\data\out\02\不存在\rep-2\rep-2.html",
            "timestamp": 100,
        },
    ]
    contacts = [
        {"姓名": "陈家瑜爸爸", "账号": "dad-1", "部门": "A/六年级2班/六年级2班家长"},
        {"姓名": "陈家瑜妈妈", "账号": "mom-1", "部门": "A/六年级2班/六年级2班家长"},
    ]
    rows, stats = mod._build_rows_from_reports(reports, contacts)
    assert len(rows) == 1
    row = rows[0]
    assert row["parent_name"] == "陈家瑜妈妈"
    assert row["student_name"] == "陈家瑜"
    assert row["class_name"] == "六年级2班"
    assert row["wecom_user_id"] == "mom-1"
    assert row["report_submission_id"] == "rep-1"
    assert row["image_path"].endswith("rep-1.png")
    assert stats["matched"] == 1
    assert stats["unmatched"] == 1


def test_build_rows_from_reports_prefers_report_or_exported_class_name_over_contact_class():
    reports = [
        {
            "id": "rep-1",
            "student_name": "张明琪",
            "class_name": "",
            "html_path": r"D:\score_reading_fresh\data\out\02\张明琪\rep-1\rep-1.html",
            "timestamp": 200,
        }
    ]
    contacts = [
        {"姓名": "张明琪/张明瑄妈妈", "账号": "mom-1", "部门": "A/六年级8班/六年级8班家长;B/一年级2班/一年级2班家长"},
    ]
    with TemporaryDirectory() as tmpdir:
        exported = Path(tmpdir) / "张明琪_六年级8班_一单元_report.png"
        exported.write_bytes(b"png")
        rows, _stats = mod._build_rows_from_reports(reports, contacts, export_images_dir=Path(tmpdir))
        assert rows[0]["class_name"] == "六年级8班"


def test_report_student_name_extracts_leading_cjk_name():
    report = {"student_name": "韩雨洋6单元背诵", "original_filename": "韩雨洋6单元背诵_06.mp3"}
    assert mod._extract_report_student_name(report) == "韩雨洋"


def test_report_driven_rows_include_required_output_columns():
    reports = [
        {
            "id": "rep-1",
            "student_name": "陈家瑜",
            "class_name": "六年级2班",
            "html_path": r"D:\score_reading_fresh\data\out\02\陈家瑜\rep-1\rep-1.html",
            "timestamp": 200,
        }
    ]
    contacts = [
        {"姓名": "陈家瑜妈妈", "账号": "mom-1", "部门": "A/六年级2班/六年级2班家长"},
    ]
    rows = mod._prepare_report_driven_rows(reports, contacts)
    assert len(rows) == 1
    row = rows[0]
    assert set(["parent_name", "student_name", "class_name", "wecom_user_id", "report_submission_id", "image_path"]).issubset(row.keys())


def test_parse_exported_report_filename_extracts_student_and_class():
    student, class_name = mod._parse_exported_report_filename("李宛迅_六年级8班_一单元_report.png")
    assert student == "李宛迅"
    assert class_name == "六年级8班"


def test_build_rows_from_reports_prefers_real_exported_png():
    reports = [
        {
            "id": "rep-1",
            "student_name": "陈家瑜",
            "class_name": "六年级2班",
            "html_path": r"D:\score_reading_fresh\data\out\02\陈家瑜\rep-1\rep-1.html",
            "timestamp": 200,
        }
    ]
    contacts = [
        {"姓名": "陈家瑜妈妈", "账号": "mom-1", "部门": "A/六年级2班/六年级2班家长"},
    ]
    with TemporaryDirectory() as tmpdir:
        exported = Path(tmpdir) / "陈家瑜_六年级2班_一单元_report.png"
        exported.write_bytes(b"png")
        rows, _stats = mod._build_rows_from_reports(reports, contacts, export_images_dir=Path(tmpdir))
        assert rows[0]["image_path"] == str(exported)


def test_build_rows_from_reports_falls_back_when_no_exported_png_exists():
    reports = [
        {
            "id": "rep-1",
            "student_name": "陈家瑜",
            "class_name": "六年级2班",
            "html_path": r"D:\score_reading_fresh\data\out\02\陈家瑜\rep-1\rep-1.html",
            "timestamp": 200,
        }
    ]
    contacts = [
        {"姓名": "陈家瑜妈妈", "账号": "mom-1", "部门": "A/六年级2班/六年级2班家长"},
    ]
    rows, _stats = mod._build_rows_from_reports(reports, contacts, export_images_dir=Path(r"Z:\definitely-missing"))
    assert rows[0]["image_path"].endswith("rep-1.png")


def test_read_contacts_xlsx_falls_back_when_pandas_missing():
    original_pd = mod.pd
    try:
        mod.pd = None
        rows = mod._read_contacts_xlsx(
            Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\data\contacts.xlsx")
        )
    finally:
        mod.pd = original_pd
    assert rows


def test_default_delta_csv_path_is_used_when_not_overridden():
    assert mod._default_delta_output_csv() == "tools/rpa_tasks.delta.csv"


def test_default_pending_csv_path_is_used_when_not_overridden():
    assert mod._default_pending_output_csv() == "tools/rpa_tasks.pending.csv"


def test_default_state_csv_path_is_used_when_not_overridden():
    assert mod._default_state_csv() == "run-logs/rpa_send_state.csv"


def test_build_delta_rows_includes_all_rows_when_previous_snapshot_missing():
    new_rows = [
        {"student_name": "陈家瑜", "class_name": "六年级2班", "report_submission_id": "rep-1"},
        {"student_name": "王小明", "class_name": "六年级2班", "report_submission_id": "rep-2"},
    ]
    assert mod._build_delta_rows(new_rows, []) == new_rows


def test_build_delta_rows_keeps_only_new_or_updated_students():
    previous_rows = [
        {"student_name": "陈家瑜", "class_name": "六年级2班", "report_submission_id": "rep-1"},
        {"student_name": "王小明", "class_name": "六年级2班", "report_submission_id": "rep-2"},
    ]
    new_rows = [
        {"student_name": "陈家瑜", "class_name": "六年级2班", "report_submission_id": "rep-1"},
        {"student_name": "王小明", "class_name": "六年级2班", "report_submission_id": "rep-3"},
        {"student_name": "李雷", "class_name": "六年级2班", "report_submission_id": "rep-4"},
    ]
    delta_rows = mod._build_delta_rows(new_rows, previous_rows)
    assert delta_rows == [
        {"student_name": "王小明", "class_name": "六年级2班", "report_submission_id": "rep-3"},
        {"student_name": "李雷", "class_name": "六年级2班", "report_submission_id": "rep-4"},
    ]


def test_build_pending_rows_excludes_handled_state_keys_only():
    full_rows = [
        {"student_name": "陈家瑜", "class_name": "六年级2班", "report_submission_id": "rep-1"},
        {"student_name": "王小明", "class_name": "六年级2班", "report_submission_id": "rep-2"},
        {"student_name": "李雷", "class_name": "六年级2班", "report_submission_id": "rep-3"},
    ]
    state_rows = [
        {"student_name": "陈家瑜", "class_name": "六年级2班", "report_submission_id": "rep-1", "status": "pasted_only"},
        {"student_name": "王小明", "class_name": "六年级2班", "report_submission_id": "rep-2", "status": "failed"},
        {"student_name": "李雷", "class_name": "六年级2班", "report_submission_id": "rep-0", "status": "sent"},
    ]
    pending_rows = mod._build_pending_rows(full_rows, state_rows)
    assert pending_rows == [
        {"student_name": "王小明", "class_name": "六年级2班", "report_submission_id": "rep-2"},
        {"student_name": "李雷", "class_name": "六年级2班", "report_submission_id": "rep-3"},
    ]

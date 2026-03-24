import importlib.util
import re
from pathlib import Path


MODULE_PATH = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\tools\wecom_rpa_gui.py")
spec = importlib.util.spec_from_file_location("wecom_rpa_gui", MODULE_PATH)
mod = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(mod)
SOURCE = MODULE_PATH.read_text(encoding="utf-8")


def test_default_task_csv_points_to_with_message():
    repo_root = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main")
    assert mod.default_tasks_csv_path(repo_root) == repo_root / "tools" / "rpa_tasks.with-message.csv"


def test_default_delta_csv_points_to_delta_file():
    repo_root = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main")
    assert mod.default_delta_csv_path(repo_root) == repo_root / "tools" / "rpa_tasks.delta.csv"


def test_default_pending_csv_points_to_pending_file():
    repo_root = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main")
    assert mod.default_pending_csv_path(repo_root) == repo_root / "tools" / "rpa_tasks.pending.csv"


def test_build_refresh_tasks_command_uses_report_driven_pipeline():
    repo_root = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main")
    cmd = mod.build_refresh_tasks_command(repo_root)
    joined = " ".join(str(part) for part in cmd)
    assert str(repo_root / "tools" / "build_rpa_tasks_from_score.py") in joined
    assert "--report-driven" in cmd
    assert str(repo_root / "data" / "contacts.xlsx") in joined
    assert str(repo_root / "tools" / "rpa_tasks.with-message.csv") in joined
    assert str(repo_root / "tools" / "rpa_tasks.delta.csv") in joined
    assert str(repo_root / "tools" / "rpa_tasks.pending.csv") in joined
    assert str(repo_root / "run-logs" / "rpa_send_state.csv") in joined
    assert "http://localhost" in joined


def test_build_backfill_state_command_uses_expected_files():
    repo_root = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main")
    cmd = mod.build_backfill_state_command(repo_root)
    joined = " ".join(str(part) for part in cmd)
    assert str(repo_root / "tools" / "backfill_rpa_send_state.py") in joined
    assert str(repo_root / "run-logs" / "rpa-results.csv") in joined
    assert str(repo_root / "tools" / "rpa_tasks.with-message.csv") in joined
    assert str(repo_root / "run-logs" / "rpa_send_state.csv") in joined


def test_parse_refresh_tasks_summary_extracts_counts():
    output = "total=20 enriched=18 no_match_or_empty=2 errors=0 delta=3 pending=11"
    assert mod.parse_refresh_tasks_summary(output) == {
        "total": 20,
        "enriched": 18,
        "no_match_or_empty": 2,
        "errors": 0,
        "delta": 3,
        "pending": 11,
    }


def test_parse_backfill_state_summary_extracts_counts():
    output = "handled_results=1 matched=1 unmatched=0 ambiguous=0 ignored_status=10 merged_total=1"
    assert mod.parse_backfill_state_summary(output) == {
        "handled_results": 1,
        "matched": 1,
        "unmatched": 0,
        "ambiguous": 0,
        "ignored_status": 10,
        "merged_total": 1,
    }


def test_finalize_refresh_result_switches_active_csv_and_status_for_pending_scope():
    repo_root = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main")
    state = {
        "csv_path": "",
        "status_text": "",
        "tasks_inspected": False,
    }

    def fake_inspect():
        state["tasks_inspected"] = True

    mod.finalize_refresh_result(
        repo_root=repo_root,
        send_scope="pending",
        summary={"total": 20, "enriched": 18, "no_match_or_empty": 2, "errors": 0, "delta": 3, "pending": 11},
        set_csv_path=lambda value: state.__setitem__("csv_path", value),
        set_status=lambda value: state.__setitem__("status_text", value),
        inspect_tasks=fake_inspect,
    )

    assert state["csv_path"] == str(repo_root / "tools" / "rpa_tasks.pending.csv")
    assert "总数20" in state["status_text"]
    assert "成功18" in state["status_text"]
    assert "未匹配2" in state["status_text"]
    assert "待发送11" in state["status_text"]
    assert state["tasks_inspected"] is True


def test_resolve_tasks_csv_for_scope_prefers_pending_by_default():
    repo_root = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main")
    assert mod.resolve_tasks_csv_for_scope(repo_root, "pending") == str(repo_root / "tools" / "rpa_tasks.pending.csv")
    assert mod.resolve_tasks_csv_for_scope(repo_root, "all") == str(repo_root / "tools" / "rpa_tasks.with-message.csv")


def test_normalize_saved_csv_path_migrates_legacy_real_csv():
    repo_root = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main")
    legacy = repo_root / "tools" / "rpa_tasks.real.csv"
    assert mod.normalize_saved_csv_path(repo_root, str(legacy)) == str(repo_root / "tools" / "rpa_tasks.with-message.csv")


class _Var:
    def __init__(self, value=""):
        self.value = value

    def set(self, value):
        self.value = value

    def get(self):
        return self.value


def test_inspect_tasks_surfaces_pending_count_in_dedicated_text(tmp_path):
    csv_path = tmp_path / "rpa_tasks.pending.csv"
    csv_path.write_text(
        "wecom_user_id,parent_name,student_name,image_path\n"
        "20-1,家长A,学生A,a.png\n"
        "20-2,家长B,学生B,b.png\n",
        encoding="utf-8-sig",
    )

    app = mod.RpaGuiApp.__new__(mod.RpaGuiApp)
    app.csv_path = _Var(str(csv_path))
    app.tasks_info = _Var()
    app.pending_info = _Var()
    app.summary_total = _Var()
    app.status_text = _Var()

    mod.RpaGuiApp.inspect_tasks(app)

    assert app.tasks_info.get() == "共 2 行，可执行 2 行"
    assert app.pending_info.get() == "待发送 2 行"
    assert app.summary_total.get() == "总数 2"
    assert app.status_text.get() == "已检查任务，共 2 行，可执行 2 行"


def test_format_backfill_summary_text_is_stable():
    assert mod.format_backfill_summary_text(
        {"matched": 1, "unmatched": 2, "ambiguous": 3, "merged_total": 4}
    ) == "历史台账：匹配 1，未匹配 2，歧义 3，累计 4"


def test_left_step_area_uses_scrollable_canvas_layout():
    assert "left_canvas = tk.Canvas(main" in SOURCE
    assert re.search(r"ttk\.Scrollbar\([^)]*orient=\"vertical\"", SOURCE)
    assert "left_canvas.create_window((0, 0), window=left" in SOURCE


def test_header_and_status_area_are_compacted():
    assert "self.header = tk.Frame(root, bg="#0f4cdb", height=68)" in SOURCE
    assert 'font=("Microsoft YaHei UI", 13, "bold")' in SOURCE
    assert 'row.grid(row=1, column=0, sticky="ew", padx=12, pady=(6, 4))' in SOURCE
    assert "pady=4" in SOURCE

from pathlib import Path
import importlib.util


REFRESH_MODULE_PATH = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_refresh.py")
SCHEMA_MODULE_PATH = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\schema.py")
RUNNER_MODULE_PATH = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\rpa_runner.py")
ROUTES_MODULE_PATH = Path(r"C:\Users\Lenovo\wecom-bulk-sender-20260222-145539\wecom-bulk-sender-main\api\app\routes.py")


def _load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(mod)
    return mod


def test_build_refresh_command_uses_report_driven_mode_and_pending_state_outputs():
    mod = _load_module(REFRESH_MODULE_PATH, "rpa_refresh")
    cmd = mod._build_refresh_command(score_api_base="http://host.docker.internal:8010")
    joined = " ".join(cmd)
    assert "build_rpa_tasks_from_score.py" in joined
    assert "--report-driven" in cmd
    assert "/data/contacts.xlsx" in joined
    assert "rpa_tasks.with-message.csv" in joined
    assert "rpa_tasks.delta.csv" in joined
    assert "rpa_tasks.pending.csv" in joined
    assert "rpa_send_state.csv" in joined
    assert "http://host.docker.internal:8010" in joined


def test_parse_refresh_summary_extracts_counts():
    mod = _load_module(REFRESH_MODULE_PATH, "rpa_refresh")
    stdout = "[OK] output: tools/rpa_tasks.with-message.csv\n[OK] delta_output: tools/rpa_tasks.delta.csv\n[OK] pending_output: tools/rpa_tasks.pending.csv\n[OK] total=5 enriched=4 no_match_or_empty=1 errors=0 delta=2 pending=3\n"
    data = mod._parse_refresh_summary(stdout)
    assert data["total"] == 5
    assert data["enriched"] == 4
    assert data["no_match_or_empty"] == 1
    assert data["errors"] == 0
    assert data["delta"] == 2
    assert data["pending"] == 3


def test_build_backfill_command_uses_expected_files():
    mod = _load_module(REFRESH_MODULE_PATH, "rpa_refresh")
    cmd = mod._build_backfill_command()
    joined = " ".join(cmd)
    assert "backfill_rpa_send_state.py" in joined
    assert "rpa-results.csv" in joined
    assert "rpa_tasks.with-message.csv" in joined
    assert "rpa_send_state.csv" in joined


def test_parse_backfill_summary_extracts_counts():
    mod = _load_module(REFRESH_MODULE_PATH, "rpa_refresh")
    stdout = "handled_results=1 matched=1 unmatched=0 ambiguous=0 ignored_status=10 merged_total=1"
    data = mod._parse_backfill_summary(stdout)
    assert data["handled_results"] == 1
    assert data["matched"] == 1
    assert data["unmatched"] == 0
    assert data["ambiguous"] == 0
    assert data["ignored_status"] == 10
    assert data["merged_total"] == 1


def test_rpa_start_request_defaults_to_pending_csv():
    mod = _load_module(SCHEMA_MODULE_PATH, "schema")
    assert mod.RpaStartRequest.model_fields["tasks_csv"].default == "tools/rpa_tasks.pending.csv"


def test_runner_resolves_default_tasks_csv_to_pending_file():
    mod = _load_module(RUNNER_MODULE_PATH, "rpa_runner")
    resolved = mod._resolve_path(None, "tools/rpa_tasks.pending.csv")
    assert resolved == mod._repo_root() / "tools" / "rpa_tasks.pending.csv"


def test_refresh_route_defaults_score_api_to_8010():
    source = ROUTES_MODULE_PATH.read_text(encoding="utf-8")
    assert 'def refresh_rpa_tasks(score_api_base: str = "http://host.docker.internal:8010")' in source

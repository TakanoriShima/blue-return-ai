"""年間処理の入口（annual / config）と source_key のテスト。すべて一時フォルダ上の架空データ。"""

from __future__ import annotations

import csv
import hashlib
import json

import pytest

from blue_return_ai import annual
from blue_return_ai.annual import AnnualError, create_run_dir, run, validate_year
from blue_return_ai.cli import UnsafeOutputDirError
from blue_return_ai.config import ConfigError, load_config
from blue_return_ai.keys import make_source_key, sales_source_key
from conftest import REPO_ROOT


def write_config(repo, data):
    path = repo / "data" / "config" / "annual.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return path


DEFAULT_CONFIG = {
    "version": 1,
    "sources": {
        "invoices": "data/invoices",
        "bank": "data/bank",
        "credit_card": "data/credit-card",
        "household": "data/household",
    },
    "decisions_dir": "data/decisions",
}


@pytest.fixture
def fake_repo(tmp_path):
    """一時フォルダに架空のリポジトリ構成と架空の資料ファイルを作る。"""
    repo = tmp_path / "repo"
    files = {
        "data/invoices/架空請求_取引先A.xlsx": b"fake excel A",
        "data/invoices/架空請求_取引先A.pdf": b"fake pdf A",
        "data/invoices/sub/架空支払明細.pdf": b"fake pdf B",
        "data/invoices/~$架空請求_取引先A.xlsx": b"lock",
        "data/invoices/メモ.txt": b"memo",
        "data/bank/架空銀行_2026.csv": b"fake bank",
        "data/bank/架空銀行_2026_重複.csv": b"fake bank",
        "data/credit-card/架空カード_202607.csv": b"fake card",
    }
    for relative, content in files.items():
        path = repo / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)
    config = write_config(repo, DEFAULT_CONFIG)
    return repo, config


def read_inventory(run_dir):
    with (run_dir / "documents_inventory.csv").open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def test_run_creates_inventory_and_manifest(fake_repo):
    repo, config = fake_repo
    result = run(2026, config, repo / "data" / "output" / "annual", repo_root=repo)
    run_dir = result["run_dir"]

    assert run_dir.parent == (repo / "data" / "output" / "annual").resolve()
    assert run_dir.name.startswith("2026_")
    rows = read_inventory(run_dir)
    assert len(rows) == 8
    by_path = {r["relative_path"]: r for r in rows}

    excel = by_path["data/invoices/架空請求_取引先A.xlsx"]
    assert excel["category"] == "invoices"
    assert excel["extension"] == ".xlsx"
    assert excel["sha256"] == hashlib.sha256(b"fake excel A").hexdigest()
    assert excel["status"] == "unresolved"
    assert excel["status_source"] == "program"
    assert by_path["data/invoices/~$架空請求_取引先A.xlsx"]["status"] == "unsupported"
    assert by_path["data/invoices/~$架空請求_取引先A.xlsx"]["note"] == "temporary_file"
    assert by_path["data/invoices/メモ.txt"]["status"] == "unsupported"
    assert by_path["data/invoices/sub/架空支払明細.pdf"]["status"] == "unresolved"
    duplicate = by_path["data/bank/架空銀行_2026_重複.csv"]
    assert duplicate["duplicate_of"] == by_path["data/bank/架空銀行_2026.csv"]["document_key"]

    manifest = json.loads((run_dir / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["year"] == 2026
    assert manifest["documents"]["total"] == 8
    assert manifest["documents"]["duplicates"] == 1
    assert manifest["documents"]["by_status"]["unresolved"] == 6
    assert manifest["documents"]["by_status"]["unsupported"] == 2
    assert manifest["sources"]["household"] == {"configured": True, "exists": False}
    # この設定には processors が無いため、資料一覧だけを作り、何も取り込まない
    assert set(manifest["processors"].values()) == {"not_configured"}
    assert manifest["processing"]["sales"]["records"] == 0


def test_manifest_contains_no_file_names(fake_repo):
    repo, config = fake_repo
    result = run(2026, config, repo / "data" / "output" / "annual", repo_root=repo)
    text = (result["run_dir"] / "manifest.json").read_text(encoding="utf-8")
    for name in ("架空請求", "架空支払明細", "架空銀行", "架空カード", "メモ"):
        assert name not in text


def test_each_run_creates_new_directory(fake_repo):
    repo, config = fake_repo
    out_root = repo / "data" / "output" / "annual"
    first = run(2026, config, out_root, repo_root=repo)["run_dir"]
    second = run(2026, config, out_root, repo_root=repo)["run_dir"]
    assert first != second
    assert first.exists() and second.exists()


def test_existing_run_directory_is_not_overwritten(tmp_path):
    out_root = tmp_path / "annual"
    create_run_dir(out_root, "2026_fixed")
    marker = out_root / "2026_fixed" / "keep.txt"
    marker.write_text("架空", encoding="utf-8")
    with pytest.raises(AnnualError):
        create_run_dir(out_root, "2026_fixed")
    assert marker.read_text(encoding="utf-8") == "架空"


def test_run_id_collision_is_rejected(fake_repo, monkeypatch):
    repo, config = fake_repo
    monkeypatch.setattr(annual, "new_run_id", lambda year: f"{year}_fixed")
    out_root = repo / "data" / "output" / "annual"
    run(2026, config, out_root, repo_root=repo)
    with pytest.raises(AnnualError):
        run(2026, config, out_root, repo_root=repo)


def test_decisions_override_status(fake_repo):
    repo, config = fake_repo
    sha = hashlib.sha256(b"fake pdf A").hexdigest()
    decisions = repo / "data" / "decisions" / "documents.csv"
    decisions.parent.mkdir(parents=True)
    decisions.write_text(f"sha256,status,note\n{sha},cross_check_only,Excel と同一の請求\n",
                         encoding="utf-8")
    result = run(2026, config, repo / "data" / "output" / "annual", repo_root=repo)
    row = next(r for r in read_inventory(result["run_dir"]) if r["sha256"] == sha)
    assert row["status"] == "cross_check_only"
    assert row["status_source"] == "human"
    assert row["note"] == "Excel と同一の請求"


@pytest.mark.parametrize(
    "content",
    ["sha256,status\nabc,manual\n", f"sha256,status\n{'0' * 64},approved\n", "hash,state\n"],
)
def test_invalid_decisions_are_rejected(fake_repo, content):
    repo, config = fake_repo
    decisions = repo / "data" / "decisions" / "documents.csv"
    decisions.parent.mkdir(parents=True)
    decisions.write_text(content, encoding="utf-8")
    with pytest.raises(AnnualError):
        run(2026, config, repo / "data" / "output" / "annual", repo_root=repo)


@pytest.mark.parametrize("year", [1999, 2100, 0, -2026, "2026", 2026.0, True, None])
def test_invalid_year_is_rejected(year):
    with pytest.raises(AnnualError):
        validate_year(year)


def test_cli_rejects_invalid_year(fake_repo, capsys):
    repo, config = fake_repo
    out_root = repo / "data" / "output" / "annual"
    assert annual.main(["--year", "1999", "--config", str(config), "--out-root", str(out_root)]) == 2
    with pytest.raises(SystemExit):
        annual.main(["--year", "abc", "--config", str(config), "--out-root", str(out_root)])
    assert not out_root.exists()


@pytest.mark.parametrize(
    "data",
    [
        {"version": 2, "sources": {"invoices": "data/invoices"}},
        {"version": 1, "sources": {}},
        {"version": 1, "sources": {"unknown": "data/x"}},
        {"version": 1, "sources": {"invoices": "sample_data/invoices"}},
        {"version": 1, "sources": {"invoices": "data/../src"}},
        {"version": 1, "sources": {"invoices": "C:/data/invoices"}},
        {"version": 1, "sources": {"invoices": 123}},
        {"version": 1, "sources": {"invoices": "data/invoices"}, "decisions_dir": "decisions"},
    ],
)
def test_invalid_config_is_rejected(tmp_path, data):
    repo = tmp_path / "repo"
    path = write_config(repo, data)
    with pytest.raises(ConfigError):
        load_config(path, repo)


def test_missing_config_is_reported(tmp_path, capsys):
    repo = tmp_path / "repo"
    with pytest.raises(ConfigError):
        load_config(repo / "data" / "config" / "annual.json", repo)


def test_config_paths_are_resolved_under_data(fake_repo):
    repo, config_path = fake_repo
    config = load_config(config_path, repo)
    assert config.sources["invoices"] == (repo / "data" / "invoices").resolve()
    assert config.decisions_dir == (repo / "data" / "decisions").resolve()
    assert len(config.sha256) == 64


def test_out_root_inside_repo_must_be_under_data(fake_repo):
    _, config = fake_repo
    target = REPO_ROOT / "sample_data" / "should_not_exist"
    with pytest.raises(UnsafeOutputDirError):
        create_run_dir(target, "2026_x")
    assert not target.exists()


def test_summary_shows_counts_only(fake_repo):
    repo, config = fake_repo
    out_root = repo / "data" / "output" / "annual"
    # CLI は REPO_ROOT 基準で設定を解決するため、テストでは repo_root 付きの run() と表示関数を使う
    result = run(2026, config, out_root, repo_root=repo)
    text = annual.format_summary(result["manifest"])
    assert "資料: 合計 8 件、同一内容の重複 1 件" in text
    assert "household: フォルダなし" in text
    for name in ("架空請求", "架空支払明細", "架空銀行", "架空カード", "メモ", ".xlsx"):
        assert name not in text


def test_example_config_is_valid():
    config = load_config(REPO_ROOT / "sample_data" / "config" / "annual.example.json", REPO_ROOT)
    assert set(config.sources) == {"invoices", "bank", "credit_card", "household"}


# ---------------------------------------------------------------- source_key

SAMPLE_HASH = hashlib.sha256(b"fake invoice").hexdigest()


def test_sales_source_key_is_stable_and_position_sensitive():
    key = sales_source_key(SAMPLE_HASH, 1)
    assert key == sales_source_key(SAMPLE_HASH, 1)
    assert key != sales_source_key(SAMPLE_HASH, 2)
    assert key.startswith("sales-1-")
    assert len(key) == len("sales-1-") + 32
    assert SAMPLE_HASH[:16] not in key


def test_source_key_differs_by_namespace_and_hides_parts():
    bank = make_source_key("bank", "2026-07-15", "98765", "架空摘要")
    card = make_source_key("card", "2026-07-15", "98765", "架空摘要")
    assert bank != card
    for part in ("2026", "98765", "架空摘要"):
        assert part not in bank


def test_source_key_known_value():
    # 実装を変えるとキーが変わり、人間の判断ファイルとの紐づけが切れるため値を固定して確認する
    expected = hashlib.sha256(
        "\x1f".join(("sales", "1", SAMPLE_HASH, "1")).encode("utf-8")).hexdigest()[:32]
    assert sales_source_key(SAMPLE_HASH) == f"sales-1-{expected}"


@pytest.mark.parametrize(
    ("call", "args"),
    [
        (sales_source_key, ("abc", 1)),
        (sales_source_key, (SAMPLE_HASH, 0)),
        (sales_source_key, (SAMPLE_HASH, True)),
        (make_source_key, ("unknown", "x")),
        (make_source_key, ("sales",)),
        (make_source_key, ("sales", "")),
    ],
)
def test_source_key_rejects_invalid_input(call, args):
    with pytest.raises(ValueError):
        call(*args)

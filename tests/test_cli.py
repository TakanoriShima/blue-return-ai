"""コマンドライン実行のテスト（完全な架空データ、出力先は tmp_path）。"""

from __future__ import annotations

import pytest

from blue_return_ai.cli import UnsafeOutputDirError, check_out_dir, main
from blue_return_ai.output import load_record_json
from conftest import REPO_ROOT


def test_cli_writes_json_and_csv(sample_invoice, tmp_path, capsys):
    out_dir = tmp_path / "out"
    assert main([str(sample_invoice), "--out-dir", str(out_dir)]) == 0

    jsons = list((out_dir / "records").glob("*.json"))
    assert len(jsons) == 1
    record = load_record_json(jsons[0])
    assert record["review_status"] == "unreviewed"

    review_files = sorted(p.name for p in (out_dir / "review").iterdir())
    assert len(review_files) == 2
    assert review_files[0].endswith("_line_items.csv")
    assert review_files[1].endswith("_records.csv")

    output = capsys.readouterr().out
    assert sample_invoice.name not in output
    assert "架空サンプル商事" not in output
    assert "61600" not in output


def test_cli_skips_same_source_hash(sample_invoice, tmp_path, capsys):
    out_dir = tmp_path / "out"
    assert main([str(sample_invoice), "--out-dir", str(out_dir)]) == 0
    assert main([str(sample_invoice), "--out-dir", str(out_dir)]) == 0
    assert len(list((out_dir / "records").glob("*.json"))) == 1
    assert "既存取込候補" in capsys.readouterr().out


def test_cli_skips_duplicate_within_one_run(sample_invoice, tmp_path):
    out_dir = tmp_path / "out"
    assert main([str(sample_invoice), str(sample_invoice), "--out-dir", str(out_dir)]) == 0
    assert len(list((out_dir / "records").glob("*.json"))) == 1


def test_cli_reports_failure_without_details(tmp_path, capsys):
    missing = tmp_path / "存在しない架空ファイル.xlsx"
    assert main([str(missing), "--out-dir", str(tmp_path / "out")]) == 1
    output = capsys.readouterr().out
    assert "FileNotFoundError" in output
    assert missing.name not in output


def test_out_dir_inside_repo_must_be_under_data():
    with pytest.raises(UnsafeOutputDirError):
        check_out_dir(REPO_ROOT / "sample_data" / "out")
    with pytest.raises(UnsafeOutputDirError):
        check_out_dir(REPO_ROOT)
    assert check_out_dir(REPO_ROOT / "data" / "output") == (REPO_ROOT / "data" / "output").resolve()


def test_cli_refuses_unsafe_out_dir(sample_invoice):
    unsafe = REPO_ROOT / "sample_data" / "should_not_exist"
    assert main([str(sample_invoice), "--out-dir", str(unsafe)]) == 2
    assert not unsafe.exists()


def test_sample_generator_refuses_data_dir_and_overwrite(sample_invoice):
    from blue_return_ai.sample_template_a import write_sample
    with pytest.raises(ValueError):
        write_sample(REPO_ROOT / "data" / "should_not_exist.xlsx")
    with pytest.raises(FileExistsError):
        write_sample(sample_invoice)


def test_cli_processes_template_b(sample_invoice_b, tmp_path, capsys):
    out_dir = tmp_path / "out"
    assert main([str(sample_invoice_b), "--template", "template_b", "--out-dir", str(out_dir)]) == 0

    jsons = list((out_dir / "records").glob("*.json"))
    assert len(jsons) == 1
    record = load_record_json(jsons[0])
    assert record["source_name"] == "template_b"
    assert record["review_status"] == "unreviewed"
    assert record["calculated_fields"] == ["net_amount"]
    assert len(list((out_dir / "review").iterdir())) == 2

    # record_id（UUID）の文字列が偶然数字列を含む場合があるため、除いてから確認する
    output = capsys.readouterr().out.replace(record["record_id"], "<record_id>")
    assert sample_invoice_b.name not in output
    assert "架空研修サービス" not in output
    for value in ("12000", "10910", "10887", "1113", "1090"):
        assert value not in output


def test_cli_template_b_with_wrong_layout_fails_without_details(sample_invoice, tmp_path, capsys):
    # テンプレートの自動判定はしない。書式が違えば読み取り失敗として扱う
    assert main([str(sample_invoice), "--template", "template_b",
                 "--out-dir", str(tmp_path / "out")]) == 1
    output = capsys.readouterr().out
    assert "SheetNotFoundError" in output
    assert sample_invoice.name not in output


def test_sample_generator_b_refuses_data_dir_and_overwrite(sample_invoice_b):
    from blue_return_ai.sample_template_b import write_sample
    with pytest.raises(ValueError):
        write_sample(REPO_ROOT / "data" / "should_not_exist.xlsx")
    with pytest.raises(FileExistsError):
        write_sample(sample_invoice_b)

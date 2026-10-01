"""JSON / 確認用 CSV 出力のテスト（完全な架空データ）。"""

from __future__ import annotations

import csv
import json

import pytest

from blue_return_ai.output import (
    LINE_ITEM_COLUMNS,
    RECORD_COLUMNS,
    escape_csv_value,
    find_records_by_source_hash,
    load_record_json,
    write_line_items_csv,
    write_record_json,
    write_review_csv,
)
from conftest import process


def read_csv(path):
    raw = path.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf"), "UTF-8 BOM 付きであること"
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def test_json_output_round_trip(sample_invoice, tmp_path):
    record = process(sample_invoice)
    path = write_record_json(record, tmp_path / "records")

    assert path.name == f"{record['record_id']}.json"
    assert load_record_json(path) == record
    text = path.read_text(encoding="utf-8")
    assert "架空サンプル商事株式会社" in text  # ensure_ascii=False
    assert json.loads(text)["revenue_date"] is None


def test_json_output_does_not_overwrite(sample_invoice, tmp_path):
    record = process(sample_invoice)
    write_record_json(record, tmp_path)
    with pytest.raises(FileExistsError):
        write_record_json(record, tmp_path)


def test_json_output_does_not_contain_source_file_name(sample_invoice, tmp_path):
    record = process(sample_invoice)
    path = write_record_json(record, tmp_path)
    assert sample_invoice.stem not in path.read_text(encoding="utf-8")


def test_find_records_by_source_hash(sample_invoice, tmp_path):
    record = process(sample_invoice)
    write_record_json(record, tmp_path)
    assert find_records_by_source_hash(tmp_path, record["source_hash"]) == [record["record_id"]]
    assert find_records_by_source_hash(tmp_path, "0" * 64) == []


def test_review_csv(sample_invoice, tmp_path):
    record = process(sample_invoice)
    path = write_review_csv([record], tmp_path / "records.csv")
    rows = read_csv(path)

    assert list(rows[0]) == RECORD_COLUMNS
    assert len(rows) == 1
    row = rows[0]
    assert row["record_id"] == record["record_id"]
    assert row["review_status"] == "unreviewed"
    assert row["gross_amount"] == "61600"
    assert row["expected_payment_amount"] == "55883"
    assert row["deductions_total"] == "0"
    assert row["service_period_month"] == "2026-02"
    assert row["service_period_from"] == ""
    assert row["human_review_status"] == ""
    assert row["human_revenue_date"] == ""


def test_review_csv_distinguishes_null_and_zero(make_invoice, tmp_path):
    zero = process(make_invoice({"E19": 0}))
    unknown = process(make_invoice({"E19": None}))
    rows = read_csv(write_review_csv([zero, unknown], tmp_path / "records.csv"))
    assert rows[0]["withholding_tax"] == "0"
    assert rows[1]["withholding_tax"] == ""
    assert "FIELD_MISSING:withholding_status" in rows[1]["warnings"]


def test_line_items_csv(sample_invoice, tmp_path):
    record = process(sample_invoice)
    rows = read_csv(write_line_items_csv([record], tmp_path / "line_items.csv"))

    assert list(rows[0]) == LINE_ITEM_COLUMNS
    assert [r["line_no"] for r in rows] == ["1", "2"]
    assert all(r["record_id"] == record["record_id"] for r in rows)
    assert rows[1]["quantity"] == "1.5"
    assert rows[1]["amount"] == "6000"


def test_csv_does_not_overwrite(sample_invoice, tmp_path):
    record = process(sample_invoice)
    path = write_review_csv([record], tmp_path / "records.csv")
    with pytest.raises(FileExistsError):
        write_review_csv([record], path)


@pytest.mark.parametrize("dangerous", ["=1+1", "+1", "-1", "@SUM(A1)", "\tX", "\rX",
                                       "=HYPERLINK(\"http://example.invalid\")"])
def test_escape_csv_value(dangerous):
    assert escape_csv_value(dangerous) == "'" + dangerous


@pytest.mark.parametrize("safe", ["架空サンプル商事株式会社", "研修", "", 100, -100, None])
def test_escape_csv_value_keeps_safe_values(safe):
    assert escape_csv_value(safe) == safe


def test_csv_injection_is_escaped_in_output(make_invoice, tmp_path):
    def text_cell_starting_with_equal(workbook):
        # 数式ではなく「= で始まる文字列」のセルとして保存する
        cell = workbook.active["B5"]
        cell.value = "=1+2"
        cell.data_type = "s"

    record = process(make_invoice({"B6": "@架空", "A10": "-研修"},
                                  workbook_hook=text_cell_starting_with_equal))
    records_rows = read_csv(write_review_csv([record], tmp_path / "records.csv"))
    item_rows = read_csv(write_line_items_csv([record], tmp_path / "line_items.csv"))

    assert records_rows[0]["customer"] == "'=1+2"
    assert records_rows[0]["description"] == "'@架空"
    assert item_rows[0]["description"] == "'-研修"
    # JSON 側の値はエスケープしない（資料の値をそのまま保持する）
    assert record["customer"] == "=1+2"

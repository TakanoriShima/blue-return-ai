"""template_a の読み取り → sales_record 生成 → 検証のテスト（完全な架空データ）。"""

from __future__ import annotations

import hashlib
import uuid

import pytest

from blue_return_ai.excel_reader import SheetNotFoundError
from conftest import codes, process


def test_sample_invoice_is_converted_to_sales_record(sample_invoice):
    record = process(sample_invoice)

    assert record["schema_version"] == "0.1"
    assert record["document_type"] == "invoice"
    assert record["source_type"] == "excel"
    assert record["source_name"] == "template_a"
    assert record["external_id"] is None
    assert record["document_date"] == "2026-02-28"
    assert record["service_period"] == {"from": None, "to": None, "month": "2026-02"}
    assert record["customer"] == "架空サンプル商事株式会社"
    assert record["description"] == "研修講師業務委託費（2026年2月分）"
    assert record["payment_due"] == "2026-03-31"
    assert record["net_amount"] == 56000
    assert record["tax_amount"] == 5600
    assert record["gross_amount"] == 61600
    assert record["tax_treatment"] == "exclusive"
    assert record["tax_rate"] == 10
    assert record["withholding_status"] == "applied"
    assert record["withholding_tax"] == 5717
    assert record["deductions"] == []
    assert record["expected_payment_amount"] == 55883
    assert record["revenue_date"] is None
    assert "payment_date" not in record
    assert record["warnings"] == []
    assert record["review_status"] == "unreviewed"


def test_line_items(sample_invoice):
    items = process(sample_invoice)["line_items"]
    assert items == [
        {"description": "研修講師業務", "quantity": "2", "unit": "日",
         "unit_price": 25000, "amount": 50000, "tax_rate": None, "tax_treatment": None},
        {"description": "研修教材作成", "quantity": "1.5", "unit": "時間",
         "unit_price": 4000, "amount": 6000, "tax_rate": None, "tax_treatment": None},
    ]


def test_quantity_is_string(sample_invoice):
    for item in process(sample_invoice)["line_items"]:
        assert isinstance(item["quantity"], str)


def test_amounts_are_int(sample_invoice):
    record = process(sample_invoice)
    for field in ("net_amount", "tax_amount", "gross_amount", "withholding_tax",
                  "expected_payment_amount"):
        assert type(record[field]) is int
    for item in record["line_items"]:
        assert type(item["unit_price"]) is int
        assert type(item["amount"]) is int


def test_source_hash_is_sha256_of_file(sample_invoice):
    record = process(sample_invoice)
    assert record["source_hash"] == hashlib.sha256(sample_invoice.read_bytes()).hexdigest()


def test_record_id_is_uuid_and_independent_of_source_hash(sample_invoice):
    first = process(sample_invoice)
    second = process(sample_invoice)

    assert uuid.UUID(first["record_id"]).version == 4
    assert first["record_id"] != second["record_id"]
    assert first["source_hash"] == second["source_hash"]
    assert first["source_hash"][:16] not in first["record_id"].replace("-", "")


def test_source_file_is_not_modified(sample_invoice):
    before = sample_invoice.read_bytes()
    mtime = sample_invoice.stat().st_mtime_ns
    process(sample_invoice)
    assert sample_invoice.read_bytes() == before
    assert sample_invoice.stat().st_mtime_ns == mtime


def test_expected_payment_from_document_is_not_marked_calculated(sample_invoice):
    record = process(sample_invoice)
    assert record["calculated_fields"] == []


def test_expected_payment_is_calculated_when_blank(make_invoice):
    record = process(make_invoice({"E20": None}))
    assert record["expected_payment_amount"] == 61600 - 5717
    assert record["calculated_fields"] == ["expected_payment_amount"]
    assert record["review_status"] == "unreviewed"


def test_withholding_zero_means_none(make_invoice):
    record = process(make_invoice({"E19": 0, "E20": None}))
    assert record["withholding_status"] == "none"
    assert record["withholding_tax"] == 0
    assert record["expected_payment_amount"] == 61600
    assert record["calculated_fields"] == ["expected_payment_amount"]


def test_withholding_blank_is_unknown_and_not_guessed(make_invoice):
    record = process(make_invoice({"E19": None, "E20": None}))
    assert record["withholding_status"] == "unknown"
    assert record["withholding_tax"] is None
    assert record["expected_payment_amount"] is None
    assert record["calculated_fields"] == []
    assert ("FIELD_MISSING", "withholding_status") in codes(record)
    assert record["review_status"] == "needs_review"


def test_missing_payment_due_is_null_and_needs_review(make_invoice):
    record = process(make_invoice({"B7": None}))
    assert record["payment_due"] is None
    assert ("FIELD_MISSING", "payment_due") in codes(record)
    assert record["review_status"] == "needs_review"


def test_missing_service_month_is_not_expanded(make_invoice):
    record = process(make_invoice({"B4": None}))
    assert record["service_period"] is None
    assert ("FIELD_MISSING", "service_period") in codes(record)


@pytest.mark.parametrize(
    ("overrides", "expected"),
    [
        ({"E18": 61000}, ("NET_TAX_GROSS_MISMATCH", "gross_amount")),
        ({"E10": 49000}, ("LINE_ITEMS_SUM_MISMATCH", "net_amount")),
        ({"E10": 49000}, ("LINE_ITEM_AMOUNT_MISMATCH", "line_items[0].amount")),
        ({"E17": 5000, "E18": 61000, "E20": 55283}, ("TAX_RATE_MISMATCH", "tax_amount")),
        ({"E20": 55000}, ("EXPECTED_PAYMENT_MISMATCH", "expected_payment_amount")),
        ({"E16": 56000.5}, ("AMOUNT_NOT_INTEGER", "net_amount")),
        ({"E19": -1}, ("AMOUNT_NEGATIVE", "withholding_tax")),
        ({"B5": None}, ("FIELD_MISSING", "customer")),
    ],
)
def test_inconsistent_invoice_needs_review(make_invoice, overrides, expected):
    record = process(make_invoice(overrides))
    assert expected in codes(record)
    assert record["review_status"] == "needs_review"


def test_due_before_document(make_invoice):
    from datetime import datetime
    record = process(make_invoice({"B7": datetime(2026, 2, 1)}))
    assert ("DUE_BEFORE_DOCUMENT", "payment_due") in codes(record)
    assert record["review_status"] == "needs_review"


def test_target_year(sample_invoice):
    assert process(sample_invoice, target_year=2026)["warnings"] == []
    record = process(sample_invoice, target_year=2025)
    assert ("DATE_OUT_OF_TARGET_YEAR", "document_date") in codes(record)


@pytest.mark.parametrize(
    ("text", "expected", "code"),
    [
        ("2026/2/28", "2026-02-28", None),
        ("2026年2月28日", "2026-02-28", None),
        ("令和8年2月28日", "2026-02-28", None),
        ("２０２６年２月２８日", "2026-02-28", None),
        ("2/28", None, "DATE_AMBIGUOUS"),
        ("2026/2/30", None, "DATE_INVALID"),
        ("近日中", None, "DATE_INVALID"),
    ],
)
def test_date_text_is_normalized(make_invoice, text, expected, code):
    record = process(make_invoice({"B3": text}))
    assert record["document_date"] == expected
    if code:
        assert (code, "document_date") in codes(record)
        assert record["review_status"] == "needs_review"


def test_formula_without_cached_value_is_null(make_invoice):
    # openpyxl で書いた数式セルはキャッシュ値を持たない
    record = process(make_invoice({"E18": "=E16+E17"}))
    assert record["gross_amount"] is None
    assert ("FORMULA_NO_CACHED_VALUE", "gross_amount") in codes(record)
    assert ("FIELD_MISSING", "gross_amount") not in codes(record)
    assert record["expected_payment_amount"] == 55883
    assert record["review_status"] == "needs_review"


def test_blank_expected_payment_formula_is_not_calculated(make_invoice):
    record = process(make_invoice({"E20": "=E18-E19"}))
    assert record["expected_payment_amount"] is None
    assert record["calculated_fields"] == []
    assert ("FORMULA_NO_CACHED_VALUE", "expected_payment_amount") in codes(record)


def test_label_mismatch_is_reported(make_invoice):
    record = process(make_invoice({"D18": "合計"}))
    assert ("TEMPLATE_LABEL_MISMATCH", "gross_amount") in codes(record)
    assert record["review_status"] == "needs_review"


def test_empty_line_items_are_null(make_invoice):
    blank = {f"{col}{row}": None for col in "ABCDE" for row in (10, 11)}
    record = process(make_invoice(blank))
    assert record["line_items"] is None
    assert ("FIELD_MISSING", "line_items") in codes(record)


def test_missing_sheet_raises(make_invoice):
    def rename(workbook):
        workbook.active.title = "別シート"
    with pytest.raises(SheetNotFoundError):
        process(make_invoice(workbook_hook=rename))


def test_warnings_contain_no_values(make_invoice):
    record = process(make_invoice({"E18": 61000, "B5": None}))
    assert record["warnings"]
    for warning in record["warnings"]:
        assert set(warning) == {"code", "field", "severity"}
        assert "61000" not in str(warning)


def test_program_never_sets_reviewed(sample_invoice, make_invoice):
    statuses = {process(sample_invoice)["review_status"],
                process(make_invoice({"E18": 1}))["review_status"]}
    assert "reviewed" not in statuses

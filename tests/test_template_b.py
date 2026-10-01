"""template_b（内税・源泉徴収型）の読み取り → sales_record 生成 → 検証のテスト（完全な架空データ）。"""

from __future__ import annotations

import hashlib
import uuid

import pytest

from blue_return_ai import template_b
from blue_return_ai.sales_record import build_sales_record, calculate_inclusive_net_amount
from conftest import codes, process


def process_b(path, **options):
    return process(path, template=template_b, **options)


def test_sample_invoice_b_is_converted_to_sales_record(sample_invoice_b):
    record = process_b(sample_invoice_b)

    assert record["schema_version"] == "0.1"
    assert record["document_type"] == "invoice"
    assert record["source_type"] == "excel"
    assert record["source_name"] == "template_b"
    assert record["external_id"] is None
    assert record["document_date"] == "2026-04-30"
    assert record["service_period"] == {"from": None, "to": None, "month": "2026-04"}
    assert record["customer"] == "架空研修サービス株式会社"
    assert record["description"] == "研修講師業務（2026年4月分）"
    assert record["payment_due"] == "2026-05-31"
    assert record["line_items"] == [
        {"description": "研修講師業務", "quantity": "4", "unit": "日",
         "unit_price": 3000, "amount": 12000, "tax_rate": None, "tax_treatment": None},
    ]
    assert record["tax_treatment"] == "inclusive"
    assert record["tax_rate"] == 10
    assert record["gross_amount"] == 12000
    assert record["tax_amount"] == 1090
    assert record["net_amount"] == 10910
    assert record["withholding_status"] == "applied"
    assert record["withholding_tax"] == 1113
    assert record["deductions"] == []
    assert record["expected_payment_amount"] == 10887
    assert record["revenue_date"] is None
    assert "payment_date" not in record


def test_net_amount_is_marked_calculated_and_expected_payment_is_from_document(sample_invoice_b):
    record = process_b(sample_invoice_b)
    assert record["calculated_fields"] == ["net_amount"]


def test_sample_invoice_b_has_no_warnings_and_is_unreviewed(sample_invoice_b):
    record = process_b(sample_invoice_b, target_year=2026)
    assert record["warnings"] == []
    assert record["review_status"] == "unreviewed"


def test_source_hash_and_record_id(sample_invoice_b):
    first = process_b(sample_invoice_b)
    second = process_b(sample_invoice_b)
    assert first["source_hash"] == hashlib.sha256(sample_invoice_b.read_bytes()).hexdigest()
    assert uuid.UUID(first["record_id"]).version == 4
    assert first["record_id"] != second["record_id"]


def test_source_file_is_not_modified(sample_invoice_b):
    before = sample_invoice_b.read_bytes()
    mtime = sample_invoice_b.stat().st_mtime_ns
    process_b(sample_invoice_b)
    assert sample_invoice_b.read_bytes() == before
    assert sample_invoice_b.stat().st_mtime_ns == mtime


@pytest.mark.parametrize(
    ("overrides", "missing_field"),
    [
        ({"F17": None}, "gross_amount"),
        ({"F18": None}, "tax_amount"),
    ],
)
def test_net_amount_is_not_calculated_when_gross_or_tax_is_missing(
        make_invoice_b, overrides, missing_field):
    record = process_b(make_invoice_b(overrides))
    assert record["net_amount"] is None
    assert "net_amount" not in record["calculated_fields"]
    assert record[missing_field] is None
    assert ("FIELD_MISSING", "net_amount") in codes(record)
    assert record["review_status"] == "needs_review"


def test_net_amount_is_not_calculated_when_tax_exceeds_gross(make_invoice_b):
    record = process_b(make_invoice_b({"F18": 13000}))
    assert record["net_amount"] is None
    assert "net_amount" not in record["calculated_fields"]
    assert record["review_status"] == "needs_review"


def test_net_amount_is_not_calculated_from_unreadable_tax(make_invoice_b):
    record = process_b(make_invoice_b({"F18": 1090.5}))
    assert record["tax_amount"] is None
    assert record["net_amount"] is None
    assert ("AMOUNT_NOT_INTEGER", "tax_amount") in codes(record)


def test_label_text_is_not_used_to_assign_fields(make_invoice_b):
    # 「小計」ラベルを変えても値の割り当てはセル定義のまま。ラベル不一致として要確認になる
    record = process_b(make_invoice_b({"E17": "税抜小計"}))
    assert record["gross_amount"] == 12000
    assert ("TEMPLATE_LABEL_MISMATCH", "gross_amount") in codes(record)
    assert record["review_status"] == "needs_review"


@pytest.mark.parametrize(
    ("overrides", "expected"),
    [
        ({"F11": 11000}, ("LINE_ITEMS_SUM_MISMATCH", "gross_amount")),
        ({"F20": 10000}, ("EXPECTED_PAYMENT_MISMATCH", "expected_payment_amount")),
        ({"F19": None}, ("FIELD_MISSING", "withholding_status")),
    ],
)
def test_inconsistent_invoice_b_needs_review(make_invoice_b, overrides, expected):
    record = process_b(make_invoice_b(overrides))
    assert expected in codes(record)
    assert record["review_status"] == "needs_review"


def test_template_a_layout_is_rejected_by_template_b(sample_invoice):
    from blue_return_ai.excel_reader import SheetNotFoundError
    with pytest.raises(SheetNotFoundError):
        process_b(sample_invoice)


def _record_with(**values) -> dict:
    extracted = {
        "document_date": "2026-04-30",
        "service_period": None,
        "customer": "架空",
        "description": None,
        "line_items": None,
        "payment_due": None,
        "net_amount": None,
        "tax_amount": 1090,
        "gross_amount": 12000,
        "tax_treatment": "inclusive",
        "tax_rate": 10,
        "withholding_status": "unknown",
        "withholding_tax": None,
        "deductions": [],
        "expected_payment_amount": None,
    }
    extracted.update(values)
    return build_sales_record(extracted=extracted, issues=[], source_hash="0" * 64,
                              source_name="test")


@pytest.mark.parametrize(
    ("values", "expected_net"),
    [
        ({}, 10910),
        ({"gross_amount": None}, None),
        ({"tax_amount": None}, None),
        ({"tax_amount": 13000}, None),
        ({"tax_treatment": "exclusive"}, None),
        ({"tax_treatment": "unknown"}, None),
        ({"tax_treatment": "none", "tax_amount": 0}, None),
        ({"net_amount": 10000}, 10000),
    ],
)
def test_calculate_inclusive_net_amount_conditions(values, expected_net):
    record = _record_with(**values)
    assert record["net_amount"] == expected_net
    calculated = expected_net is not None and "net_amount" not in values
    assert ("net_amount" in record["calculated_fields"]) == calculated


def test_calculate_inclusive_net_amount_does_not_overwrite_document_value():
    record = {"tax_treatment": "inclusive", "net_amount": 10000,
              "gross_amount": 12000, "tax_amount": 1090}
    assert calculate_inclusive_net_amount(record) is None


def test_net_amount_with_issue_is_not_calculated():
    issue = {"code": "FORMULA_NO_CACHED_VALUE", "field": "net_amount", "severity": "warning"}
    extracted = _record_with()  # 生成済みレコードから抽出値相当を作る
    extracted["net_amount"] = None
    record = build_sales_record(extracted=extracted, issues=[issue], source_hash="0" * 64,
                                source_name="test")
    assert record["net_amount"] is None
    assert "net_amount" not in record["calculated_fields"]

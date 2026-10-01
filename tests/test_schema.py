"""書式の異なる template_a / template_b が、同じ共通 sales_record スキーマへ正規化されることのテスト。"""

from __future__ import annotations

from blue_return_ai import template_a, template_b
from conftest import process

EXPECTED_KEYS = [
    "schema_version", "record_id", "document_type", "source_type", "source_name",
    "source_hash", "external_id", "document_date", "service_period", "customer",
    "description", "line_items", "net_amount", "tax_amount", "gross_amount",
    "tax_treatment", "tax_rate", "withholding_status", "withholding_tax", "deductions",
    "expected_payment_amount", "payment_due", "revenue_date", "review_status",
    "warnings", "calculated_fields",
]
LINE_ITEM_KEYS = ["description", "quantity", "unit", "unit_price", "amount",
                  "tax_rate", "tax_treatment"]
INT_FIELDS = ["net_amount", "tax_amount", "gross_amount", "withholding_tax",
              "expected_payment_amount"]


def test_both_templates_share_the_same_schema(sample_invoice, sample_invoice_b):
    record_a = process(sample_invoice, template=template_a)
    record_b = process(sample_invoice_b, template=template_b)

    for record in (record_a, record_b):
        assert list(record) == EXPECTED_KEYS
        assert record["schema_version"] == "0.1"
        assert set(record["service_period"]) == {"from", "to", "month"}
        for item in record["line_items"]:
            assert list(item) == LINE_ITEM_KEYS
            assert isinstance(item["quantity"], str)
        for field in INT_FIELDS:
            assert type(record[field]) is int
        assert record["warnings"] == []
        assert record["review_status"] == "unreviewed"
        assert record["revenue_date"] is None

    assert record_a["source_name"] == "template_a"
    assert record_b["source_name"] == "template_b"


def test_different_documents_keep_their_own_meaning(sample_invoice, sample_invoice_b):
    record_a = process(sample_invoice, template=template_a)
    record_b = process(sample_invoice_b, template=template_b)

    # template_a: 外税。税抜金額は資料の値
    assert record_a["tax_treatment"] == "exclusive"
    assert "net_amount" not in record_a["calculated_fields"]
    # template_b: 内税。税抜金額は資料に無く、税込総額 − 内消費税額で計算した値
    assert record_b["tax_treatment"] == "inclusive"
    assert record_b["calculated_fields"] == ["net_amount"]

    for record in (record_a, record_b):
        assert record["net_amount"] + record["tax_amount"] == record["gross_amount"]
        assert (record["gross_amount"] - record["withholding_tax"]
                == record["expected_payment_amount"])

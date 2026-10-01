"""共通 sales_record の生成（docs/data_model.md §3、schema_version 0.1）。"""

from __future__ import annotations

import uuid

SCHEMA_VERSION = "0.1"


def build_sales_record(
    *,
    extracted: dict,
    issues: list[dict],
    source_hash: str,
    source_name: str,
    document_type: str = "invoice",
    source_type: str = "excel",
) -> dict:
    """アダプターの抽出値から sales_record を組み立てる。

    - record_id は UUID v4 で生成し、source_hash からは生成しない。
    - external_id / revenue_date は MVP-1 では常に null。
    - review_status は検証（validation.validate）で決めるため、ここでは仮の値を入れる。
    """
    record = {
        "schema_version": SCHEMA_VERSION,
        "record_id": str(uuid.uuid4()),
        "document_type": document_type,
        "source_type": source_type,
        "source_name": source_name,
        "source_hash": source_hash,
        "external_id": None,
        "document_date": extracted["document_date"],
        "service_period": extracted["service_period"],
        "customer": extracted["customer"],
        "description": extracted["description"],
        "line_items": extracted["line_items"],
        "net_amount": extracted["net_amount"],
        "tax_amount": extracted["tax_amount"],
        "gross_amount": extracted["gross_amount"],
        "tax_treatment": extracted["tax_treatment"],
        "tax_rate": extracted["tax_rate"],
        "withholding_status": extracted["withholding_status"],
        "withholding_tax": extracted["withholding_tax"],
        "deductions": extracted["deductions"],
        "expected_payment_amount": extracted["expected_payment_amount"],
        "payment_due": extracted["payment_due"],
        "revenue_date": None,
        "review_status": "needs_review",
        "warnings": list(issues),
        "calculated_fields": [],
    }

    expected_cell_has_issue = any(i["field"] == "expected_payment_amount" for i in issues)
    if record["expected_payment_amount"] is None and not expected_cell_has_issue:
        calculated = calculate_expected_payment(record)
        if calculated is not None:
            record["expected_payment_amount"] = calculated
            record["calculated_fields"].append("expected_payment_amount")

    return record


def calculate_expected_payment(record: dict) -> int | None:
    """§5.5 の計算式。必要な値がすべて確定している場合のみ計算する。"""
    gross = record["gross_amount"]
    status = record["withholding_status"]
    withholding = record["withholding_tax"]
    deductions = record["deductions"]

    if gross is None or deductions is None:
        return None
    if status not in ("none", "applied") or withholding is None:
        return None
    if any(d.get("amount") is None for d in deductions):
        return None
    return gross - withholding - sum(d["amount"] for d in deductions)

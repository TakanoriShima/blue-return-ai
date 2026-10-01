"""sales_record の共通検証（docs/data_model.md §5.6、§6、§7、§8）。

検証は値を修正しない。問題は warnings に記録し、review_status を決める。
プログラムは review_status に reviewed を設定しない。
"""

from __future__ import annotations

import copy
from datetime import date
from decimal import ROUND_CEILING, ROUND_FLOOR, ROUND_HALF_UP, Decimal

from .issues import make_issue

DEFAULT_MAX_AMOUNT = 10_000_000
DEFAULT_MAX_DUE_DAYS = 120

REQUIRED_FIELDS = ("document_date", "customer", "gross_amount")
AMOUNT_FIELDS = ("net_amount", "tax_amount", "gross_amount", "withholding_tax",
                 "expected_payment_amount")
LINE_ITEM_FIELDS = ("description", "quantity", "unit", "unit_price", "amount")

VALID_WITHHOLDING = {
    # (withholding_status, withholding_tax が None か) → 追加条件
    ("none", False): lambda v: v == 0,
    ("applied", False): lambda v: v > 0,
    ("applied", True): lambda v: True,
    ("unknown", True): lambda v: True,
}


def validate(
    record: dict,
    *,
    target_year: int | None = None,
    max_amount: int = DEFAULT_MAX_AMOUNT,
    max_due_days: int = DEFAULT_MAX_DUE_DAYS,
) -> dict:
    """検証済みの sales_record（コピー）を返す。"""
    result = copy.deepcopy(record)
    issues: list[dict] = result["warnings"]

    def add(code: str, field: str | None) -> None:
        issue = make_issue(code, field)
        if issue not in issues:
            issues.append(issue)

    _check_missing(result, issues, add)
    _check_amounts(result, add, max_amount)
    _check_line_items(result, add)
    _check_withholding(result, add)
    _check_expected_payment(result, add)
    _check_dates(result, add, target_year, max_due_days)

    has_missing_required = any(result[f] is None for f in REQUIRED_FIELDS)
    result["review_status"] = "needs_review" if (issues or has_missing_required) else "unreviewed"
    return result


def _check_missing(record: dict, issues: list[dict], add) -> None:
    """取得できなかった項目に FIELD_MISSING を付ける（既に別の警告がある項目は除く）。"""
    fields_with_issue = {i["field"] for i in issues}

    def missing(field: str) -> None:
        if field not in fields_with_issue:
            add("FIELD_MISSING", field)

    for field in ("document_date", "service_period", "customer", "description",
                  "line_items", "payment_due", "net_amount", "gross_amount",
                  "deductions", "expected_payment_amount"):
        if record[field] is None:
            missing(field)

    treatment = record["tax_treatment"]
    if treatment == "unknown":
        missing("tax_treatment")
    if record["tax_amount"] is None and treatment != "inclusive":
        missing("tax_amount")
    if record["tax_rate"] is None and treatment == "exclusive":
        missing("tax_rate")

    if record["withholding_status"] == "unknown":
        missing("withholding_status")
    elif record["withholding_status"] == "applied" and record["withholding_tax"] is None:
        missing("withholding_tax")

    for index, item in enumerate(record["line_items"] or []):
        for key in LINE_ITEM_FIELDS:
            if item[key] is None:
                missing(f"line_items[{index}].{key}")


def _check_amounts(record: dict, add, max_amount: int) -> None:
    for field in AMOUNT_FIELDS:
        value = record[field]
        if value is None:
            continue
        if value < 0:
            add("AMOUNT_NEGATIVE", field)
        if value > max_amount:
            add("AMOUNT_UNUSUALLY_LARGE", field)
    for index, deduction in enumerate(record["deductions"] or []):
        if deduction.get("amount") is not None and deduction["amount"] < 0:
            add("AMOUNT_NEGATIVE", f"deductions[{index}].amount")

    net, tax, gross = record["net_amount"], record["tax_amount"], record["gross_amount"]
    treatment = record["tax_treatment"]

    # §5.2 tax_treatment と金額の関係
    if treatment == "none":
        if tax is not None and tax != 0:
            add("NET_TAX_GROSS_MISMATCH", "tax_amount")
        if net is not None and gross is not None and net != gross:
            add("NET_TAX_GROSS_MISMATCH", "gross_amount")
    elif None not in (net, tax, gross) and net + tax != gross:
        add("NET_TAX_GROSS_MISMATCH", "gross_amount")

    # 消費税率との整合。端数処理はいずれかと一致すれば可とする
    rate = record["tax_rate"]
    if treatment == "exclusive" and None not in (net, tax, rate):
        exact = Decimal(net) * Decimal(rate) / Decimal(100)
        candidates = {int(exact.to_integral_value(rounding=r))
                      for r in (ROUND_FLOOR, ROUND_HALF_UP, ROUND_CEILING)}
        if tax not in candidates:
            add("TAX_RATE_MISMATCH", "tax_amount")


def _check_line_items(record: dict, add) -> None:
    items = record["line_items"]
    if not items:
        return

    for index, item in enumerate(items):
        for key in ("unit_price", "amount"):
            if item[key] is not None and item[key] < 0:
                add("AMOUNT_NEGATIVE", f"line_items[{index}].{key}")
        quantity, unit_price, amount = item["quantity"], item["unit_price"], item["amount"]
        if None not in (quantity, unit_price, amount):
            if Decimal(quantity) * Decimal(unit_price) != Decimal(amount):
                add("LINE_ITEM_AMOUNT_MISMATCH", f"line_items[{index}].amount")

    amounts = [item["amount"] for item in items]
    if None in amounts:
        return
    treatment = record["tax_treatment"]
    target_field = {"exclusive": "net_amount", "none": "net_amount",
                    "inclusive": "gross_amount"}.get(treatment)
    if target_field is None or record[target_field] is None:
        return
    if sum(amounts) != record[target_field]:
        add("LINE_ITEMS_SUM_MISMATCH", target_field)


def _check_withholding(record: dict, add) -> None:
    status, value = record["withholding_status"], record["withholding_tax"]
    rule = VALID_WITHHOLDING.get((status, value is None))
    if rule is None or not rule(value):
        add("WITHHOLDING_INCONSISTENT", "withholding_tax")


def _check_expected_payment(record: dict, add) -> None:
    """資料に記載された振込予定額を §5.5 の計算式で検証する。"""
    if "expected_payment_amount" in record["calculated_fields"]:
        return
    expected = record["expected_payment_amount"]
    gross, withholding = record["gross_amount"], record["withholding_tax"]
    deductions = record["deductions"]
    if None in (expected, gross, withholding, deductions):
        return
    if record["withholding_status"] not in ("none", "applied"):
        return
    if any(d.get("amount") is None for d in deductions):
        return
    if gross - withholding - sum(d["amount"] for d in deductions) != expected:
        add("EXPECTED_PAYMENT_MISMATCH", "expected_payment_amount")


def _check_dates(record: dict, add, target_year: int | None, max_due_days: int) -> None:
    document_date = _to_date(record["document_date"])
    payment_due = _to_date(record["payment_due"])

    if document_date and target_year is not None and document_date.year != target_year:
        add("DATE_OUT_OF_TARGET_YEAR", "document_date")
    if document_date and payment_due:
        if payment_due < document_date:
            add("DUE_BEFORE_DOCUMENT", "payment_due")
        elif (payment_due - document_date).days > max_due_days:
            add("DUE_TOO_FAR", "payment_due")

    period = record["service_period"]
    if period:
        start, end = _to_date(period.get("from")), _to_date(period.get("to"))
        if start and end and start > end:
            add("PERIOD_REVERSED", "service_period")


def _to_date(value: str | None) -> date | None:
    return date.fromisoformat(value) if value else None

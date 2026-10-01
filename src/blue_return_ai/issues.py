"""検証・抽出で検出した警告（docs/data_model.md §8.2）。

警告には実データの値を含めない。コード・項目名・重大度のみで表現する。
"""

from __future__ import annotations

ERROR = "error"
WARNING = "warning"

SEVERITY: dict[str, str] = {
    # 抽出（§11）
    "FORMULA_NO_CACHED_VALUE": WARNING,
    "TEMPLATE_LABEL_MISMATCH": ERROR,
    "QUANTITY_INVALID": WARNING,
    # 欠落（§8.2）
    "FIELD_MISSING": WARNING,
    # 金額（§5.6）
    "AMOUNT_NOT_INTEGER": ERROR,
    "AMOUNT_NEGATIVE": ERROR,
    "NET_TAX_GROSS_MISMATCH": ERROR,
    "LINE_ITEMS_SUM_MISMATCH": ERROR,
    "LINE_ITEM_AMOUNT_MISMATCH": WARNING,
    "TAX_RATE_MISMATCH": WARNING,
    "WITHHOLDING_INCONSISTENT": ERROR,
    "EXPECTED_PAYMENT_MISMATCH": ERROR,
    "AMOUNT_UNUSUALLY_LARGE": WARNING,
    # 日付（§7）
    "DATE_INVALID": ERROR,
    "DATE_AMBIGUOUS": WARNING,
    "DATE_OUT_OF_TARGET_YEAR": WARNING,
    "DUE_BEFORE_DOCUMENT": ERROR,
    "DUE_TOO_FAR": WARNING,
    "PERIOD_REVERSED": ERROR,
}


def make_issue(code: str, field: str | None) -> dict:
    """警告を 1 件作る。未定義のコードは受け付けない。"""
    return {"code": code, "field": field, "severity": SEVERITY[code]}

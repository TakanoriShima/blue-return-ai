"""テンプレート template_a の定義と項目割り当て（docs/data_model.md §10）。

template_a は、開発・公開用の完全な架空の請求書書式である。
どのセルがどの項目に当たるかは、ラベル名から推測せず、このテンプレート定義で固定する。
"""

from __future__ import annotations

from .cell_parsing import (
    parse_amount,
    parse_date,
    parse_month,
    parse_quantity,
    parse_text,
)
from .excel_reader import ExcelSource, read_excel
from .issues import make_issue

SOURCE_NAME = "template_a"
SHEET_NAME = "請求書"

# 書式の確認用ラベル。期待どおりでなければ、別書式の可能性があるため警告する
LABELS = {
    "A3": "請求日",
    "A4": "業務対象月",
    "A5": "取引先",
    "A6": "件名",
    "A7": "支払期限",
    "A9": "明細",
    "B9": "数量",
    "C9": "単位",
    "D9": "単価",
    "E9": "金額",
    "D16": "税抜金額",
    "D17": "消費税額",
    "D18": "税込総額",
    "D19": "源泉徴収税額",
    "D20": "振込予定額",
}

# ラベルと対応する値セル（ラベル確認の警告をどの項目に付けるかにも使う）
LABEL_FIELDS = {
    "A3": "document_date",
    "A4": "service_period",
    "A5": "customer",
    "A6": "description",
    "A7": "payment_due",
    "A9": "line_items",
    "B9": "line_items",
    "C9": "line_items",
    "D9": "line_items",
    "E9": "line_items",
    "D16": "net_amount",
    "D17": "tax_amount",
    "D18": "gross_amount",
    "D19": "withholding_tax",
    "D20": "expected_payment_amount",
}

CELLS = {
    "document_date": "B3",
    "service_month": "B4",
    "customer": "B5",
    "description": "B6",
    "payment_due": "B7",
    "net_amount": "E16",
    "tax_amount": "E17",
    "gross_amount": "E18",
    "withholding_tax": "E19",
    "expected_payment_amount": "E20",
}

LINE_ITEM_ROWS = range(10, 15)
LINE_ITEM_COLUMNS = {
    "description": "A",
    "quantity": "B",
    "unit": "C",
    "unit_price": "D",
    "amount": "E",
}

# テンプレート定義者（人間）が確認済みとして扱う既定値（§10）
# template_a は外税・税率 10% の書式であり、源泉徴収以外の控除欄を持たない
DEFAULTS = {
    "tax_treatment": "exclusive",
    "tax_rate": 10,
    "deductions": [],
}


def extract(source: ExcelSource) -> tuple[dict, list[dict]]:
    """template_a のセルから sales_record 用の値を取り出す。

    戻り値の dict には、資料から取得した値（取得できなければ None）のみを入れる。
    計算による補完は行わない。
    """
    issues: list[dict] = []

    def take(parser, coordinate: str, field: str):
        value, code = parser(source.cell(coordinate))
        if code is not None:
            issues.append(make_issue(code, field))
        return value

    for coordinate, expected in LABELS.items():
        actual, _ = parse_text(source.cell(coordinate))
        if actual != expected:
            issues.append(make_issue("TEMPLATE_LABEL_MISMATCH", LABEL_FIELDS[coordinate]))

    month = take(parse_month, CELLS["service_month"], "service_period")
    withholding_tax = take(parse_amount, CELLS["withholding_tax"], "withholding_tax")
    withholding_cell_has_issue = any(i["field"] == "withholding_tax" for i in issues)

    extracted = {
        "document_date": take(parse_date, CELLS["document_date"], "document_date"),
        "service_period": (
            {"from": None, "to": None, "month": month} if month is not None else None
        ),
        "customer": take(parse_text, CELLS["customer"], "customer"),
        "description": take(parse_text, CELLS["description"], "description"),
        "line_items": _extract_line_items(source, take),
        "payment_due": take(parse_date, CELLS["payment_due"], "payment_due"),
        "net_amount": take(parse_amount, CELLS["net_amount"], "net_amount"),
        "tax_amount": take(parse_amount, CELLS["tax_amount"], "tax_amount"),
        "gross_amount": take(parse_amount, CELLS["gross_amount"], "gross_amount"),
        "tax_treatment": DEFAULTS["tax_treatment"],
        "tax_rate": DEFAULTS["tax_rate"],
        "withholding_status": _withholding_status(withholding_tax, withholding_cell_has_issue),
        "withholding_tax": withholding_tax,
        "deductions": list(DEFAULTS["deductions"]),
        "expected_payment_amount": take(
            parse_amount, CELLS["expected_payment_amount"], "expected_payment_amount"
        ),
    }
    return extracted, issues


def _withholding_status(withholding_tax: int | None, cell_has_issue: bool) -> str:
    """源泉徴収税額欄の記載から状態を決める（§5.3）。

    - 欄が空、または読み取れない: unknown（記載が無いだけで none としない）
    - 0 と明記: none
    - それ以外の金額: applied
    """
    if withholding_tax is None or cell_has_issue:
        return "unknown"
    if withholding_tax == 0:
        return "none"
    return "applied"


def _extract_line_items(source: ExcelSource, take) -> list[dict] | None:
    items: list[dict] = []
    for row in LINE_ITEM_ROWS:
        coordinates = {key: f"{col}{row}" for key, col in LINE_ITEM_COLUMNS.items()}
        if all(source.cell(c).value is None and not source.cell(c).is_formula
               for c in coordinates.values()):
            continue
        index = len(items)
        items.append({
            "description": take(parse_text, coordinates["description"], f"line_items[{index}].description"),
            "quantity": take(parse_quantity, coordinates["quantity"], f"line_items[{index}].quantity"),
            "unit": take(parse_text, coordinates["unit"], f"line_items[{index}].unit"),
            "unit_price": take(parse_amount, coordinates["unit_price"], f"line_items[{index}].unit_price"),
            "amount": take(parse_amount, coordinates["amount"], f"line_items[{index}].amount"),
            "tax_rate": None,
            "tax_treatment": None,
        })
    # 明細欄が空の場合は「明細なしと確定」ではなく「取得できない」とする
    return items or None


def read(path) -> tuple[str, dict, list[dict]]:
    """Excel ファイルを読み、(source_hash, 抽出値, 抽出時の警告) を返す。"""
    source = read_excel(path, SHEET_NAME)
    extracted, issues = extract(source)
    return source.source_hash, extracted, issues

"""テンプレート template_b の定義と項目割り当て（docs/data_model.md §10）。

template_b は、開発・公開用の完全な架空の請求書書式である（内税・源泉徴収型）。
template_a とはレイアウトも金額の意味も異なる。

- 帳票上の「小計」は税込総額（gross_amount）を意味する。
- 「（内消費税）」は内税の消費税額（tax_amount）。
- 税抜金額の欄は無い（net_amount は資料から取得しない）。

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

SOURCE_NAME = "template_b"
SHEET_NAME = "請求明細"

# 書式の確認用ラベル。期待どおりでなければ、別書式の可能性があるため警告する
LABELS = {
    "A1": "御請求書",
    "A3": "宛先",
    "A4": "発行日",
    "A5": "対象期間",
    "A6": "件名",
    "A7": "お支払期限",
    "B10": "品目",
    "C10": "数量",
    "D10": "単位",
    "E10": "単価",
    "F10": "金額",
    "E17": "小計",
    "E18": "（内消費税）",
    "E19": "源泉徴収税",
    "E20": "お振込金額",
}

# ラベルと対応する項目（ラベル確認の警告をどの項目に付けるかに使う）
LABEL_FIELDS = {
    "A1": None,
    "A3": "customer",
    "A4": "document_date",
    "A5": "service_period",
    "A6": "description",
    "A7": "payment_due",
    "B10": "line_items",
    "C10": "line_items",
    "D10": "line_items",
    "E10": "line_items",
    "F10": "line_items",
    "E17": "gross_amount",
    "E18": "tax_amount",
    "E19": "withholding_tax",
    "E20": "expected_payment_amount",
}

CELLS = {
    "customer": "B3",
    "document_date": "B4",
    "service_month": "B5",
    "description": "B6",
    "payment_due": "B7",
    "gross_amount": "F17",  # 帳票上の「小計」。この書式では税込総額
    "tax_amount": "F18",  # 帳票上の「（内消費税）」
    "withholding_tax": "F19",
    "expected_payment_amount": "F20",
}

LINE_ITEM_ROWS = range(11, 16)
LINE_ITEM_COLUMNS = {
    "description": "B",
    "quantity": "C",
    "unit": "D",
    "unit_price": "E",
    "amount": "F",
}

# テンプレート定義者（人間）が確認済みとして扱う既定値（§10）
# template_b は内税・税率 10% の書式であり、源泉徴収以外の控除欄を持たない
DEFAULTS = {
    "tax_treatment": "inclusive",
    "tax_rate": 10,
    "deductions": [],
}


def extract(source: ExcelSource) -> tuple[dict, list[dict]]:
    """template_b のセルから sales_record 用の値を取り出す。

    戻り値の dict には、資料から取得した値（取得できなければ None）のみを入れる。
    net_amount は資料に記載が無いため None とし、計算はしない（共通処理 sales_record で扱う）。
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
        "net_amount": None,
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
    """源泉徴収税欄の記載から状態を決める（§5.3。template_a と同じ規則）。

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

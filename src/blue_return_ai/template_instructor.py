"""テンプレート template_instructor（シート「講師」の請求書 Excel）の定義と項目割り当て（docs/data_model.md §17.1.2）。

template_c（請求書作成サービス系の書式）とは構造が大きく異なるため、独立したアダプターとする。
人間が構造調査で確認した構造だけを使い、公開コードには実在の氏名・会社名・金額を含めない。

- 請求日：「請求日」ラベルの右側の値（日付セル・文字列・シリアル値）。
- 明細見出し：「商品名」「数量」「単位」「単価(円)」「税率」「金額(円)」が同じ行に各 1 つ、確認済みの列
  （A / D / E / F / G / H）にあること。列が違う・不足・重複する場合は推測せず TEMPLATE_LABEL_MISMATCH。
- 明細：見出しの次の行から、最初の集計欄・税率別内訳のラベルの直前まで。全項目が空の行は飛ばす。
- 集計欄：「小計」「消費税」「源泉税額」「合計」が見出しより下に各 1 つ、F 列にあること。値はラベルの右側。
- 割り当て：小計 → net_amount、消費税 → tax_amount、源泉税額 → withholding_tax。
  「合計」は源泉徴収控除後の支払額かどうかを構造だけでは確定できないため、共通モデルに割り当てず
  document_total（確認用）として保持する。gross_amount も計算で作らない（null）。
- tax_treatment：税率別内訳（「税率別内訳」「税抜金額」「消費税額」）がそろっている場合だけ exclusive、
  それ以外は unknown。税率別内訳の数値は使わない。
- 請求先・件名・支払期限・業務期間は欄を確認していないため null（推測しない）。
"""

from __future__ import annotations

import re
import unicodedata
from decimal import Decimal, InvalidOperation

from .cell_parsing import (
    normalize_label,
    parse_amount,
    parse_excel_date,
    parse_quantity,
    parse_text,
)
from .excel_labels import column_of, label_cells, labels_in_row, row_of, value_right_of
from .excel_reader import CellValue, ExcelSource, read_excel
from .issues import make_issue

SOURCE_NAME = "template_instructor"
SHEET_NAME = "講師"
PATTERN_NAME = "instructor"

DOCUMENT_DATE_LABEL = "請求日"

# 明細見出し {項目: (ラベル, 確認済みの列)}
LINE_HEADERS = {
    "description": ("商品名", "A"),
    "quantity": ("数量", "D"),
    "unit": ("単位", "E"),
    "unit_price": ("単価(円)", "F"),
    "tax_rate": ("税率", "G"),
    "amount": ("金額(円)", "H"),
}

# 集計欄 {項目: ラベル}。ラベルは確認済みの列（F）にあること
SUMMARY_LABEL_COLUMN = "F"
SUMMARY_LABELS = {
    "net_amount": "小計",
    "tax_amount": "消費税",
    "withholding_tax": "源泉税額",
    "document_total": "合計",  # 意味（控除後の支払額か）が確定しないため共通モデルに割り当てない
}

# 外税の帳票であることを構造で確認するためのラベル（数値は使わない）
TAX_BREAKDOWN_LABELS = ("税率別内訳", "税抜金額", "消費税額")

_PERCENT = re.compile(r"^(\d+(?:\.\d+)?)%$")


def extract(source: ExcelSource) -> tuple[dict, list[dict]]:
    """「講師」シートのセルから sales_record 用の値を取り出す（資料の値のみ。計算・推測はしない）。"""
    issues: list[dict] = []

    def take(parser, coordinate: str | None, field_name: str):
        if coordinate is None:
            return None
        value, code = parser(source.cell(coordinate))
        if code is not None:
            issues.append(make_issue(code, field_name))
        return value

    def mismatch(field_name: str | None) -> None:
        issues.append(make_issue("TEMPLATE_LABEL_MISMATCH", field_name))

    labels = label_cells(source)

    # 請求日
    document_date = None
    date_cells = labels.get(normalize_label(DOCUMENT_DATE_LABEL), [])
    if len(date_cells) == 1:
        document_date = take(lambda cell: parse_excel_date(cell, source.epoch),
                             value_right_of(source, date_cells[0]), "document_date")
    else:
        mismatch("document_date")

    # 明細見出し
    header_row, columns = _find_line_header(labels)
    structure_ok = header_row is not None

    # 集計欄
    summary_cells: dict[str, str] = {}
    if structure_ok:
        for key, label in SUMMARY_LABELS.items():
            cells = [c for c in labels.get(normalize_label(label), []) if row_of(c) > header_row]
            if len(cells) != 1 or column_of(cells[0]) != SUMMARY_LABEL_COLUMN:
                structure_ok = False
                break
            summary_cells[key] = cells[0]

    breakdown_rows = [
        row_of(c) for label in TAX_BREAKDOWN_LABELS
        for c in labels.get(normalize_label(label), [])
        if header_row is not None and row_of(c) > header_row
    ]
    has_breakdown = all(
        any(header_row is not None and row_of(c) > header_row
            for c in labels.get(normalize_label(label), []))
        for label in TAX_BREAKDOWN_LABELS)

    amounts: dict[str, int | None] = {}
    line_items = None
    if structure_ok:
        for key, cell in summary_cells.items():
            amounts[key] = take(parse_amount, value_right_of(source, cell), key)
        last_row = min([row_of(c) for c in summary_cells.values()] + breakdown_rows) - 1
        line_items = _extract_line_items(source, take, columns, header_row + 1, last_row)
    else:
        mismatch(None)
        mismatch("line_items")

    withholding_tax = amounts.get("withholding_tax")
    withholding_has_issue = any(i["field"] == "withholding_tax" for i in issues)
    if not structure_ok or withholding_tax is None or withholding_has_issue:
        withholding_status = "unknown"
    elif withholding_tax == 0:
        withholding_status = "none"
    else:
        withholding_status = "applied"

    extracted = {
        "document_date": document_date,
        "service_period": None,
        "customer": None,  # 請求先の位置は未確認
        "description": None,
        "line_items": line_items,
        "payment_due": None,
        "net_amount": amounts.get("net_amount"),
        "tax_amount": amounts.get("tax_amount"),
        "gross_amount": None,  # 「小計＋消費税」を計算して税込総額にしない
        "tax_treatment": "exclusive" if structure_ok and has_breakdown else "unknown",
        "tax_rate": None,  # 税率は明細ごとの値として保持する
        "withholding_status": withholding_status,
        "withholding_tax": withholding_tax,
        "deductions": None,  # 控除の有無を確認できる欄が無い
        "expected_payment_amount": None,  # 「合計」の意味が確定しないため割り当てない
        # 共通モデルには割り当てない確認用の値
        "template_pattern": PATTERN_NAME if structure_ok else f"{PATTERN_NAME}_unidentified",
        "document_total": amounts.get("document_total"),
        "document_subtotal": None,
        "withholding_base": None,
    }
    return extracted, issues


def _find_line_header(labels: dict[str, list[str]]) -> tuple[int | None, dict[str, str]]:
    """明細見出しの行と列を確認する。確認できなければ (None, {})。"""
    anchors = labels.get(normalize_label(LINE_HEADERS["description"][0]), [])
    if len(anchors) != 1:
        return None, {}
    header_row = row_of(anchors[0])
    in_row = labels_in_row(labels, header_row)
    expected = {normalize_label(label): (key, column) for key, (label, column) in LINE_HEADERS.items()}
    names = [label for label, _ in in_row]
    if sorted(names) != sorted(expected):
        return None, {}
    columns = {}
    for label, cell in in_row:
        key, column = expected[label]
        if column_of(cell) != column:
            return None, {}
        columns[key] = column
    return header_row, columns


def parse_tax_rate(cell: CellValue):
    """明細の税率を、パーセントの整数（例：10）にする。「10%」・10・0.1（%表示）に対応。"""
    value = cell.value
    if value is None:
        return None, None
    if isinstance(value, bool):
        return None, "FIELD_MISSING"
    try:
        if isinstance(value, (int, float)):
            number = Decimal(str(value))
            number = number * 100 if 0 < number < 1 else number
        elif isinstance(value, str):
            match = _PERCENT.match(unicodedata.normalize("NFKC", value).strip())
            if not match:
                return None, "FIELD_MISSING"
            number = Decimal(match.group(1))
        else:
            return None, "FIELD_MISSING"
    except InvalidOperation:
        return None, "FIELD_MISSING"
    if number != number.to_integral_value() or not 0 <= number <= 100:
        return None, "FIELD_MISSING"
    return int(number), None


def _extract_line_items(source: ExcelSource, take, columns: dict[str, str],
                        first_row: int, last_row: int) -> list[dict] | None:
    items: list[dict] = []
    for row in range(first_row, last_row + 1):
        coordinates = {key: f"{col}{row}" for key, col in columns.items()}
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
            "tax_rate": take(parse_tax_rate, coordinates["tax_rate"], f"line_items[{index}].tax_rate"),
            "tax_treatment": None,
        })
    # 明細欄が空の場合は「明細なしと確定」ではなく「取得できない」とする
    return items or None


def read(path) -> tuple[str, dict, list[dict]]:
    """Excel ファイルを読み、(source_hash, 抽出値, 抽出時の警告) を返す。"""
    source = read_excel(path, SHEET_NAME)
    extracted, issues = extract(source)
    return source.source_hash, extracted, issues

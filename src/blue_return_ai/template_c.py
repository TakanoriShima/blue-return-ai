"""テンプレート template_c の定義と項目割り当て（docs/data_model.md §10、§17.1）。

template_c は、本人が実際に使用している請求書 Excel の書式群を、人間が構造調査で確認した結果に基づいて
定義したもの。公開コードには実在のサービス名・取引先名・個人名・金額を含めない。

同じシート名でも、明細見出しと集計欄のラベル構成が異なる 5 つのパターン（A〜E）が確認されている。
パターンはファイル名・取引先名・金額では判定せず、ラベルの集合だけで判定する。

- ラベルの比較は NFKC 正規化と空白除去の後の完全一致。似たラベル（例：「金額」と「合計金額(内税)」）を同じ意味として扱わない。
- 明細見出し「品番•品名」の行にあるラベルの集合が、いずれかのパターンの明細見出しと完全に一致すること。
- 見出しより下にある「既知の集計ラベル」の集合が、そのパターンの集計ラベルと完全に一致し、各 1 つであること。
- 一致しない・曖昧な場合は推測せず TEMPLATE_LABEL_MISMATCH とし、金額を割り当てない。
- 明細は「見出しの次の行」から「最も上の集計ラベルの直前」まで。全項目が空の行は飛ばす。値は各見出しと同じ列。
- 集計欄・請求日・ご請求金額の値は、ラベルと同じ行でラベルより右にある「最も近い空でないセル」。
  右に値が無ければ null（値が型に合わなければ AMOUNT_NOT_INTEGER / DATE_INVALID）。
- 帳票に無い・意味の確認できない項目は null とし、0・none・推測値で埋めない。
"""

from __future__ import annotations

from dataclasses import dataclass

from .cell_parsing import (
    normalize_label,
    parse_amount,
    parse_excel_date,
    parse_quantity,
    parse_text,
)
from .excel_labels import label_cells, labels_in_row, row_of, column_of, value_right_of
from .excel_reader import ExcelSource, read_excel
from .issues import make_issue

SOURCE_NAME = "template_c"
SHEET_NAME = "misoca_invoice"

LINE_ANCHOR_LABEL = "品 番 • 品 名"
DOCUMENT_DATE_LABEL = "請求日："
DOCUMENT_TOTAL_LABEL = "ご請求金額"


@dataclass(frozen=True)
class Pattern:
    """確認済みの帳票パターン。line_headers / summary は {項目: ラベル}。

    summary の項目は共通モデルの金額項目か、共通モデルに割り当てない確認用の値
    （document_subtotal：意味未確定の小計、withholding_base：源泉徴収の対象額）。
    """
    name: str
    line_headers: dict[str, str]
    summary: dict[str, str]
    tax_treatment: str
    tax_rate: int | None
    unit: str | None
    reads_customer: bool = False
    title: str | None = None
    withholding_label: bool = False
    # 振込額（差引支払額）の欄があり控除欄が無い書式。deductions を []（控除なし）とする（テンプレート定義）
    payment_label: bool = False
    # 帳票に欄が無く、推測で埋めない項目（FIELD_MISSING を付ける）
    unavailable: tuple[str, ...] = ()


PATTERNS: tuple[Pattern, ...] = (
    # Pattern A: 外税＋源泉徴収
    Pattern(
        name="A",
        line_headers={"description": "品 番 • 品 名", "quantity": "数 量（日）",
                      "unit_price": "単 価", "amount": "金 額"},
        summary={"net_amount": "小計（税抜き）", "tax_amount": "消費税（10%）",
                 "gross_amount": "税込合計金額", "withholding_tax": "源泉徴収税額",
                 "expected_payment_amount": "振込依頼金額"},
        tax_treatment="exclusive", tax_rate=10, unit="日",
        reads_customer=True, title="請　求　書", withholding_label=True, payment_label=True,
    ),
    # Pattern B: 内税＋源泉徴収（税抜金額・内税額の欄は無い。報酬額(源泉対象)は共通モデルに割り当てない）
    Pattern(
        name="B",
        line_headers={"description": "品 番 • 品 名", "quantity": "数 量（時間）",
                      "unit_price": "単 価", "amount": "合計金額(内税)"},
        summary={"gross_amount": "合計金額(税込み)", "withholding_base": "報酬額(源泉対象)",
                 "withholding_tax": "源泉徴収税(10.21%)",
                 "expected_payment_amount": "振込金額(控除後支払額)"},
        tax_treatment="inclusive", tax_rate=None, unit="時間", withholding_label=True,
        payment_label=True, unavailable=("tax_amount", "tax_rate"),
    ),
    # Pattern C: 外税、源泉徴収欄の確認なし
    Pattern(
        name="C",
        line_headers={"description": "品 番 • 品 名", "quantity": "数 量（時間）",
                      "unit_price": "単 価", "amount": "金 額"},
        summary={"net_amount": "小計", "tax_amount": "消費税（10%）", "gross_amount": "合計金額"},
        tax_treatment="exclusive", tax_rate=10, unit="時間",
    ),
    # Pattern D: 小計のみ。税・源泉徴収の意味は未確定
    Pattern(
        name="D",
        line_headers={"description": "品 番 • 品 名", "quantity": "数 量",
                      "unit_price": "単 価", "amount": "金額"},
        summary={"document_subtotal": "小計"},
        tax_treatment="unknown", tax_rate=None, unit=None,
    ),
    # Pattern E: 内税の明細＋税抜小計・内消費税・源泉徴収・差引支払額（明細見出しは B と同じだが別パターン）
    Pattern(
        name="E",
        line_headers={"description": "品 番 • 品 名", "quantity": "数 量（時間）",
                      "unit_price": "単 価", "amount": "合計金額(内税)"},
        summary={"gross_amount": "合計(税込)", "net_amount": "小計(税抜)",
                 "tax_amount": "消費税(内税10%)", "withholding_tax": "源泉徴収税額",
                 "expected_payment_amount": "合計（差引支払額）"},
        tax_treatment="inclusive", tax_rate=10, unit="時間", withholding_label=True,
        payment_label=True,
    ),
)

KNOWN_SUMMARY_LABELS = {normalize_label(label) for p in PATTERNS for label in p.summary.values()}


def identify_pattern(source: ExcelSource) -> tuple[Pattern | None, int | None, dict[str, str]]:
    """(パターン, 明細見出しの行, {項目: 集計ラベルのセル}) を返す。判定できなければパターンは None。"""
    labels = label_cells(source)
    anchors = labels.get(normalize_label(LINE_ANCHOR_LABEL), [])
    if len(anchors) != 1:
        return None, None, {}
    header_row = row_of(anchors[0])
    header_labels = [label for label, _ in labels_in_row(labels, header_row)]
    if len(header_labels) != len(set(header_labels)):
        return None, header_row, {}

    below = {label: [c for c in cells if row_of(c) > header_row]
             for label, cells in labels.items() if label in KNOWN_SUMMARY_LABELS}
    below = {label: cells for label, cells in below.items() if cells}

    # 明細見出しが同じパターン（B と E）があるため、見出しと集計欄の両方が完全に一致するものを集める
    matches = []
    for pattern in PATTERNS:
        if set(header_labels) != {normalize_label(v) for v in pattern.line_headers.values()}:
            continue
        expected = {normalize_label(label): key for key, label in pattern.summary.items()}
        if set(below) == set(expected) and all(len(cells) == 1 for cells in below.values()):
            matches.append((pattern, {expected[label]: cells[0] for label, cells in below.items()}))
    if len(matches) != 1:
        return None, header_row, {}
    pattern, summary_cells = matches[0]
    return pattern, header_row, summary_cells


def extract(source: ExcelSource) -> tuple[dict, list[dict]]:
    """template_c のセルから sales_record 用の値を取り出す（資料の値のみ。計算・推測はしない）。

    戻り値には、sales_record の項目に加えて template_pattern / document_total / document_subtotal /
    withholding_base を含める（共通モデルには割り当てず、確認用に保持する値）。
    """
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

    def labelled_value(label: str, field_name: str, parser, *, required: bool):
        cells = labels.get(normalize_label(label), [])
        if len(cells) != 1:
            if cells or required:
                mismatch(field_name)
            return None
        return take(parser, value_right_of(source, cells[0]), field_name)

    pattern, header_row, summary_cells = identify_pattern(source)
    if pattern is None:
        mismatch(None)

    document_date = labelled_value(
        DOCUMENT_DATE_LABEL, "document_date",
        lambda cell: parse_excel_date(cell, source.epoch), required=True)
    document_total = labelled_value(DOCUMENT_TOTAL_LABEL, "document_total", parse_amount,
                                    required=False)

    amounts: dict[str, int | None] = {}
    for key, label_cell in summary_cells.items():
        amounts[key] = take(parse_amount, value_right_of(source, label_cell), key)

    line_items = None
    customer = None
    if pattern is not None:
        if pattern.title is not None:
            title, _ = parse_text(source.cell("B1"))
            if title is None or normalize_label(title) != normalize_label(pattern.title):
                mismatch(None)
        if pattern.reads_customer:
            customer = take(parse_text, "B3", "customer")
        columns = {}
        for key, label in pattern.line_headers.items():
            cell = next(c for c in labels[normalize_label(label)] if row_of(c) == header_row)
            columns[key] = column_of(cell)
        last_row = min(row_of(c) for c in summary_cells.values()) - 1
        line_items = _extract_line_items(source, take, columns, pattern.unit, header_row + 1, last_row)
        for field_name in pattern.unavailable:
            issues.append(make_issue("FIELD_MISSING", field_name))
    else:
        mismatch("line_items")

    withholding_tax = amounts.get("withholding_tax")
    withholding_has_issue = any(i["field"] == "withholding_tax" for i in issues)
    has_withholding_label = pattern is not None and pattern.withholding_label

    extracted = {
        "document_date": document_date,
        "service_period": None,  # この書式群には業務対象期間の欄が無い
        "customer": customer,  # 請求先の位置はパターン A でのみ確認済み
        "description": None,  # 件名の欄は無い（明細から推測しない）
        "line_items": line_items,
        "payment_due": None,  # 支払期限の欄は無い
        "net_amount": amounts.get("net_amount"),
        "tax_amount": amounts.get("tax_amount"),
        "gross_amount": amounts.get("gross_amount"),
        "tax_treatment": pattern.tax_treatment if pattern else "unknown",
        "tax_rate": pattern.tax_rate if pattern else None,
        "withholding_status": _withholding_status(
            withholding_tax, withholding_has_issue, has_withholding_label),
        "withholding_tax": withholding_tax,
        "deductions": [] if pattern is not None and pattern.payment_label else None,
        "expected_payment_amount": amounts.get("expected_payment_amount"),
        # 共通モデルには割り当てない確認用の値
        "template_pattern": pattern.name if pattern else None,
        "document_total": document_total,
        "document_subtotal": amounts.get("document_subtotal"),
        "withholding_base": amounts.get("withholding_base"),
    }
    return extracted, issues


def _withholding_status(withholding_tax: int | None, cell_has_issue: bool, label_found: bool) -> str:
    """源泉徴収欄の記載から状態を決める（§5.3）。欄が無い・空・読めない場合は unknown。"""
    if not label_found or withholding_tax is None or cell_has_issue:
        return "unknown"
    if withholding_tax == 0:
        return "none"
    return "applied"


def _extract_line_items(source: ExcelSource, take, columns: dict[str, str], unit: str | None,
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
            "unit": unit,
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

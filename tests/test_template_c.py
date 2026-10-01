"""template_c（パターン A〜D をラベル集合で判定する書式）のテスト。すべて完全な架空データ。"""

from __future__ import annotations

from datetime import datetime

import pytest
from openpyxl.utils.datetime import to_excel

from blue_return_ai import template_c
from blue_return_ai.cell_parsing import parse_date, parse_excel_date
from blue_return_ai.excel_reader import CellValue
from blue_return_ai.sample_template_c import build_workbook, write_sample
from conftest import REPO_ROOT, codes, process

SAMPLE_C = REPO_ROOT / "sample_data" / "invoices" / "template_c_sample.xlsx"

# この書式群に欄が無い項目（推測で補完しないため FIELD_MISSING になる）
NO_FIELD = {("FIELD_MISSING", "service_period"), ("FIELD_MISSING", "description"),
            ("FIELD_MISSING", "payment_due")}


def process_c(path, **options):
    return process(path, template=template_c, **options)


def extract_c(path):
    return template_c.read(path)[1]


@pytest.fixture
def make_c(tmp_path):
    counter = {"n": 0}

    def _make(**kwargs):
        counter["n"] += 1
        path = tmp_path / f"invoice_c_{counter['n']}.xlsx"
        hook = kwargs.pop("hook", None)
        if hook is None:
            return write_sample(path, **kwargs)
        workbook = build_workbook(**kwargs)
        hook(workbook.active)
        workbook.save(path)
        return path

    return _make


def clear_label(label):
    def hook(sheet):
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == label:
                    cell.value = None
    return hook


# ---------------------------------------------------------------- Pattern A

def test_pattern_a_is_read_as_before(make_c):
    record = process_c(SAMPLE_C, target_year=2026)
    assert extract_c(SAMPLE_C)["template_pattern"] == "A"
    assert record["document_date"] == "2026-05-31"
    assert record["customer"] == "架空サンプル研修株式会社"
    assert (record["net_amount"], record["tax_amount"], record["gross_amount"]) == (70000, 7000, 77000)
    assert record["tax_treatment"] == "exclusive"
    assert record["tax_rate"] == 10
    assert record["withholding_status"] == "applied"
    assert record["withholding_tax"] == 7147
    assert record["expected_payment_amount"] == 69853
    assert record["deductions"] == []
    assert record["calculated_fields"] == []
    assert record["revenue_date"] is None
    assert [i["unit"] for i in record["line_items"]] == ["日", "日"]
    assert len(record["line_items"]) == 2
    assert codes(record) == NO_FIELD


def test_pattern_a_document_total_is_kept_as_reference(make_c):
    extracted = extract_c(make_c(pattern="A"))
    assert extracted["document_total"] == 77000
    assert "架空 請求者" not in str(extracted)  # 氏名欄は読まない


@pytest.mark.parametrize("count", [1, 3, 8])
def test_pattern_a_row_count_differences(make_c, count):
    items = [(f"架空業務{i}", 1, 1000, 1000) for i in range(count)]
    total = 1000 * count
    record = process_c(make_c(pattern="A", items=items, summary={
        "net_amount": total, "tax_amount": total // 10, "gross_amount": total + total // 10,
        "withholding_tax": 0, "expected_payment_amount": total + total // 10}))
    assert len(record["line_items"]) == count
    assert record["withholding_status"] == "none"
    assert record["withholding_tax"] == 0
    assert codes(record) == NO_FIELD


def test_blank_rows_before_summary_are_skipped(make_c):
    record = process_c(make_c(pattern="A", blank_rows_after_items=2))
    assert len(record["line_items"]) == 2
    assert codes(record) == NO_FIELD


@pytest.mark.parametrize(
    ("summary", "expected"),
    [
        ({"gross_amount": 77001}, ("NET_TAX_GROSS_MISMATCH", "gross_amount")),
        ({"expected_payment_amount": 70000}, ("EXPECTED_PAYMENT_MISMATCH", "expected_payment_amount")),
        ({"tax_amount": 6000, "gross_amount": 76000, "expected_payment_amount": 68853},
         ("TAX_RATE_MISMATCH", "tax_amount")),
    ],
)
def test_pattern_a_amount_consistency(make_c, summary, expected):
    record = process_c(make_c(pattern="A", summary=summary))
    assert expected in codes(record)


def test_pattern_a_formula_without_cache(make_c):
    def formula(sheet):
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == "税込合計金額":
                    sheet[f"L{cell.row}"] = "=1+1"
    record = process_c(make_c(pattern="A", hook=formula))
    assert record["gross_amount"] is None
    assert ("FORMULA_NO_CACHED_VALUE", "gross_amount") in codes(record)


# ---------------------------------------------------------------- Pattern B

def test_pattern_b_reads_only_safe_fields(make_c):
    path = make_c(pattern="B")
    extracted = extract_c(path)
    record = process_c(path)
    assert extracted["template_pattern"] == "B"
    assert record["tax_treatment"] == "inclusive"
    assert record["gross_amount"] == 33000
    assert record["withholding_status"] == "applied"
    assert record["withholding_tax"] == 3063
    assert record["expected_payment_amount"] == 29937
    # 税抜金額・内税額・税率は帳票に無いため推測しない
    assert record["net_amount"] is None
    assert record["tax_amount"] is None
    assert record["tax_rate"] is None
    assert record["calculated_fields"] == []
    assert {("FIELD_MISSING", "net_amount"), ("FIELD_MISSING", "tax_amount"),
            ("FIELD_MISSING", "tax_rate")} <= codes(record)
    # 報酬額(源泉対象) は共通モデルに割り当てず、確認用の値として保持する
    assert extracted["withholding_base"] == 30000
    # 請求先の位置はパターン A でのみ確認済み
    assert record["customer"] is None
    assert record["line_items"][0]["unit"] == "時間"
    assert record["line_items"][0]["amount"] == 33000
    assert ("LINE_ITEMS_SUM_MISMATCH", "gross_amount") not in codes(record)
    assert ("EXPECTED_PAYMENT_MISMATCH", "expected_payment_amount") not in codes(record)
    assert not any(c == "TEMPLATE_LABEL_MISMATCH" for c, _ in codes(record))


# ---------------------------------------------------------------- Pattern C

def test_pattern_c(make_c):
    path = make_c(pattern="C")
    record = process_c(path)
    assert extract_c(path)["template_pattern"] == "C"
    assert (record["net_amount"], record["tax_amount"], record["gross_amount"]) == (10000, 1000, 11000)
    assert record["tax_treatment"] == "exclusive"
    assert record["tax_rate"] == 10
    # 源泉徴収欄・支払額欄が無いため「なし」「0」と推測しない
    assert record["withholding_status"] == "unknown"
    assert record["withholding_tax"] is None
    assert record["expected_payment_amount"] is None
    assert record["deductions"] is None
    assert record["calculated_fields"] == []
    assert ("FIELD_MISSING", "withholding_status") in codes(record)
    assert ("FIELD_MISSING", "expected_payment_amount") in codes(record)
    assert record["line_items"][0]["unit"] == "時間"
    assert not any(c == "TEMPLATE_LABEL_MISMATCH" for c, _ in codes(record))


# ---------------------------------------------------------------- Pattern D

def test_pattern_d_without_guessing(make_c):
    path = make_c(pattern="D")
    extracted = extract_c(path)
    record = process_c(path)
    assert extracted["template_pattern"] == "D"
    assert len(record["line_items"]) == 1
    assert record["line_items"][0]["amount"] == 8000
    assert record["line_items"][0]["unit"] is None
    # 小計の税務上の意味は推測しない
    for key in ("net_amount", "tax_amount", "gross_amount", "withholding_tax",
                "expected_payment_amount"):
        assert record[key] is None
    assert record["tax_treatment"] == "unknown"
    assert record["withholding_status"] == "unknown"
    assert extracted["document_subtotal"] == 8000
    assert extracted["document_total"] == 8000
    assert {("FIELD_MISSING", "gross_amount"), ("FIELD_MISSING", "tax_treatment"),
            ("FIELD_MISSING", "withholding_status")} <= codes(record)
    assert not any(c == "TEMPLATE_LABEL_MISMATCH" for c, _ in codes(record))
    assert record["review_status"] == "needs_review"


# ---------------------------------------------------------------- パターン判定

@pytest.mark.parametrize(
    ("pattern", "header_cell", "label"),
    [
        ("A", "J13", "数 量（時間）"),       # A の見出しに C の数量 → 集計欄が合わない
        ("C", "N13", "合計金額(内税)"),      # C の見出しに B の金額 → 集計欄が合わない
        ("D", "J13", "数量（日）"),          # D の見出しの数量だけ A → どのパターンでもない
        ("A", "N13", "合計金額(内税)"),      # 似たラベルを同じ意味として扱わない
    ],
)
def test_mixed_labels_are_not_guessed(make_c, pattern, header_cell, label):
    path = make_c(pattern=pattern, overrides={header_cell: label})
    extracted = extract_c(path)
    record = process_c(path)
    assert extracted["template_pattern"] is None
    assert ("TEMPLATE_LABEL_MISMATCH", None) in codes(record)
    assert ("TEMPLATE_LABEL_MISMATCH", "line_items") in codes(record)
    assert record["gross_amount"] is None
    assert record["line_items"] is None


@pytest.mark.parametrize("pattern", ["A", "B", "C"])
def test_missing_summary_label_is_not_guessed(make_c, pattern):
    label = next(iter(template_c.PATTERNS[["A", "B", "C"].index(pattern)].summary.values()))
    path = make_c(pattern=pattern, hook=clear_label(label))
    assert extract_c(path)["template_pattern"] is None
    assert ("TEMPLATE_LABEL_MISMATCH", None) in codes(process_c(path))


def test_duplicate_summary_label_is_not_guessed(make_c):
    path = make_c(pattern="A", overrides={"J30": "税込合計金額"})
    assert extract_c(path)["template_pattern"] is None
    assert process_c(path)["gross_amount"] is None


def test_extra_known_summary_label_makes_pattern_ambiguous(make_c):
    # D の帳票に、意味の確定した別パターンの集計ラベルが混在している
    path = make_c(pattern="D", overrides={"J30": "消費税（10%）"})
    assert extract_c(path)["template_pattern"] is None


def test_missing_or_duplicate_line_anchor(make_c):
    path = make_c(pattern="A", hook=clear_label("品 番 • 品 名"))
    assert extract_c(path)["template_pattern"] is None
    path = make_c(pattern="A", overrides={"B30": "品 番 • 品 名"})
    assert extract_c(path)["template_pattern"] is None


def test_label_spacing_is_normalized_but_characters_are_not(make_c):
    path = make_c(pattern="A", overrides={"J13": "数量（日）", "B13": "品　番　•　品　名"})
    assert extract_c(path)["template_pattern"] == "A"
    path = make_c(pattern="A", overrides={"B13": "品番・品名"})
    assert extract_c(path)["template_pattern"] is None


def test_pattern_a_title_mismatch(make_c):
    record = process_c(make_c(pattern="A", overrides={"B1": "見積書"}))
    assert ("TEMPLATE_LABEL_MISMATCH", None) in codes(record)


def test_missing_date_label(make_c):
    record = process_c(make_c(pattern="C", hook=clear_label("請求日：")))
    assert ("TEMPLATE_LABEL_MISMATCH", "document_date") in codes(record)
    assert record["document_date"] is None


def test_other_sheet_name_is_not_template_c(tmp_path):
    from openpyxl import Workbook

    from blue_return_ai.excel_reader import SheetNotFoundError
    workbook = Workbook()
    workbook.active.title = "架空の別書式"
    path = tmp_path / "other.xlsx"
    workbook.save(path)
    with pytest.raises(SheetNotFoundError):
        template_c.read(path)


# ---------------------------------------------------------------- Excel の日付

@pytest.mark.parametrize("pattern", ["A", "B", "C", "D"])
def test_excel_serial_date_is_converted(make_c, pattern):
    serial = to_excel(datetime(2026, 4, 30))
    assert isinstance(serial, (int, float))
    record = process_c(make_c(pattern=pattern, document_date=serial))
    assert record["document_date"] == "2026-04-30"
    assert ("DATE_INVALID", "document_date") not in codes(record)


@pytest.mark.parametrize(
    ("value", "expected"),
    [(datetime(2026, 4, 30), "2026-04-30"), ("2026/4/30", "2026-04-30"),
     ("令和8年4月30日", "2026-04-30")],
)
def test_existing_date_formats_still_work(make_c, value, expected):
    record = process_c(make_c(pattern="A", document_date=value))
    assert record["document_date"] == expected


@pytest.mark.parametrize("value", [0, -1, 3_000_000, "2026/2/30", "近日"])
def test_invalid_serial_or_date_is_date_invalid(make_c, value):
    record = process_c(make_c(pattern="A", document_date=value))
    assert record["document_date"] is None
    assert ("DATE_INVALID", "document_date") in codes(record)


def test_parse_excel_date_uses_workbook_epoch():
    from openpyxl.utils.datetime import MAC_EPOCH
    serial_1904 = to_excel(datetime(2026, 4, 30), epoch=MAC_EPOCH)
    cell = CellValue(value=serial_1904, is_formula=False)
    assert parse_excel_date(cell, MAC_EPOCH) == ("2026-04-30", None)
    assert parse_excel_date(cell)[0] != "2026-04-30"  # 1900 年基準で読むと別の日付になる
    assert parse_excel_date(CellValue(value=True, is_formula=False)) == (None, "DATE_INVALID")
    assert parse_excel_date(CellValue(value=float("nan"), is_formula=False)) == (None, "DATE_INVALID")
    assert parse_excel_date(CellValue(value=None, is_formula=True)) == (None, "FORMULA_NO_CACHED_VALUE")


def test_parse_date_for_other_templates_is_unchanged():
    # template_a / template_b が使う parse_date は、数値をシリアル値として解釈しない（従来どおり）
    assert parse_date(CellValue(value=46142, is_formula=False)) == (None, "DATE_INVALID")


# ---------------------------------------------------------------- Pattern E

def test_pattern_e_is_identified_and_read(make_c):
    path = make_c(pattern="E")
    extracted = extract_c(path)
    record = process_c(path, target_year=2026)
    assert extracted["template_pattern"] == "E"
    assert (record["net_amount"], record["tax_amount"], record["gross_amount"]) == (10000, 1000, 11000)
    assert record["tax_treatment"] == "inclusive"
    assert record["tax_rate"] == 10
    assert record["withholding_status"] == "applied"
    assert record["withholding_tax"] == 1021
    assert record["expected_payment_amount"] == 9979
    assert record["deductions"] == []
    assert record["calculated_fields"] == []  # 税抜金額は帳票の値（計算していない）
    assert record["line_items"] == [
        {"description": "架空講義", "quantity": "4", "unit": "時間", "unit_price": 2750,
         "amount": 11000, "tax_rate": None, "tax_treatment": None}]
    assert record["customer"] is None
    assert codes(record) == NO_FIELD | {("FIELD_MISSING", "customer")}


def test_pattern_e_multiple_items_and_blank_rows(make_c):
    items = [(f"架空講義{i}", 1, 1100, 1100) for i in range(3)]
    record = process_c(make_c(pattern="E", items=items, blank_rows_after_items=2, summary={
        "gross_amount": 3300, "net_amount": 3000, "tax_amount": 300,
        "withholding_tax": 306, "expected_payment_amount": 2994}))
    assert len(record["line_items"]) == 3
    assert codes(record) == NO_FIELD | {("FIELD_MISSING", "customer")}


def test_pattern_e_withholding_zero_is_none(make_c):
    record = process_c(make_c(pattern="E", summary={"withholding_tax": 0,
                                                    "expected_payment_amount": 11000}))
    assert record["withholding_status"] == "none"
    assert record["withholding_tax"] == 0


@pytest.mark.parametrize(
    ("summary", "expected"),
    [
        ({"tax_amount": 900}, ("NET_TAX_GROSS_MISMATCH", "gross_amount")),
        ({"gross_amount": 12000, "net_amount": 11000, "expected_payment_amount": 10979},
         ("LINE_ITEMS_SUM_MISMATCH", "gross_amount")),
        ({"expected_payment_amount": 10000}, ("EXPECTED_PAYMENT_MISMATCH", "expected_payment_amount")),
    ],
)
def test_pattern_e_consistency_uses_existing_warnings(make_c, summary, expected):
    record = process_c(make_c(pattern="E", summary=summary))
    assert expected in codes(record)
    assert record["review_status"] == "needs_review"


@pytest.mark.parametrize("label", ["小計(税抜)", "消費税(内税10%)", "合計（差引支払額）"])
def test_pattern_e_missing_label_is_not_guessed(make_c, label):
    path = make_c(pattern="E", hook=clear_label(label))
    assert extract_c(path)["template_pattern"] is None
    assert process_c(path)["gross_amount"] is None


def test_pattern_e_duplicate_label_is_not_guessed(make_c):
    path = make_c(pattern="E", overrides={"J30": "合計(税込)"})
    assert extract_c(path)["template_pattern"] is None


def test_pattern_b_and_e_are_not_confused(make_c):
    # 明細見出しは同じ。集計欄のラベル構成で区別する
    assert extract_c(make_c(pattern="B"))["template_pattern"] == "B"
    assert extract_c(make_c(pattern="E"))["template_pattern"] == "E"
    mixed = make_c(pattern="E", overrides={"J30": "振込金額(控除後支払額)"})
    assert extract_c(mixed)["template_pattern"] is None
    mixed = make_c(pattern="B", overrides={"J30": "小計(税抜)"})
    assert extract_c(mixed)["template_pattern"] is None


def test_pattern_e_with_other_pattern_summary_label_is_not_guessed(make_c):
    path = make_c(pattern="E", overrides={"J30": "小計（税抜き）"})
    record = process_c(path)
    assert extract_c(path)["template_pattern"] is None
    assert ("TEMPLATE_LABEL_MISMATCH", None) in codes(record)
    assert record["net_amount"] is None

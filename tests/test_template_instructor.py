"""template_instructor（シート「講師」）のテスト。すべて完全な架空データ。"""

from __future__ import annotations

from datetime import datetime

import pytest
from openpyxl.utils.datetime import to_excel

from blue_return_ai import template_c, template_instructor
from blue_return_ai.excel_reader import CellValue, SheetNotFoundError
from blue_return_ai.sample_template_instructor import build_workbook, write_sample
from blue_return_ai.template_instructor import parse_tax_rate
from conftest import codes, process

# 欄の位置を確認していない、または意味が確定しないため値を入れない項目
NOT_ASSIGNED = {("FIELD_MISSING", f) for f in (
    "customer", "description", "payment_due", "service_period", "gross_amount",
    "expected_payment_amount", "deductions", "tax_rate")}


def process_i(path, **options):
    return process(path, template=template_instructor, **options)


def extract_i(path):
    return template_instructor.read(path)[1]


@pytest.fixture
def make_i(tmp_path):
    counter = {"n": 0}

    def _make(**kwargs):
        counter["n"] += 1
        path = tmp_path / f"instructor_{counter['n']}.xlsx"
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


def test_instructor_sheet_is_read(make_i):
    path = make_i()
    extracted = extract_i(path)
    record = process_i(path, target_year=2026)
    assert extracted["template_pattern"] == "instructor"
    assert record["source_name"] == "template_instructor"
    assert record["document_date"] == "2026-06-30"
    assert record["net_amount"] == 13000
    assert record["tax_amount"] == 1300
    assert record["withholding_tax"] == 1327
    assert record["withholding_status"] == "applied"
    assert record["tax_treatment"] == "exclusive"
    assert record["revenue_date"] is None
    assert record["line_items"] == [
        {"description": "架空講座A", "quantity": "2", "unit": "時間", "unit_price": 5000,
         "amount": 10000, "tax_rate": 10, "tax_treatment": None},
        {"description": "架空講座B", "quantity": "1", "unit": "回", "unit_price": 3000,
         "amount": 3000, "tax_rate": 10, "tax_treatment": None},
    ]
    assert codes(record) == NOT_ASSIGNED
    assert record["review_status"] == "needs_review"


def test_meaning_unconfirmed_values_are_not_guessed(make_i):
    path = make_i()
    extracted = extract_i(path)
    record = process_i(path)
    # 「合計」は控除後の支払額と確定できないため、確認用の値としてのみ保持する
    assert extracted["document_total"] == 12973
    assert record["expected_payment_amount"] is None
    # 小計＋消費税を計算して税込総額にしない
    assert record["gross_amount"] is None
    assert record["calculated_fields"] == []
    assert record["customer"] is None
    assert record["deductions"] is None
    # 氏名・振込先などの欄は読まない
    assert "架空 講師" not in str(extracted)
    assert "架空銀行" not in str(extracted)


def test_single_item_and_blank_rows(make_i):
    record = process_i(make_i(items=[("架空単発講座", 1, "回", 8000, "10%", 8000)],
                              summary={"net_amount": 8000, "tax_amount": 800,
                                       "withholding_tax": 0, "document_total": 8800},
                              blank_rows_after_items=2))
    assert len(record["line_items"]) == 1
    assert record["withholding_status"] == "none"
    assert record["withholding_tax"] == 0
    assert codes(record) == NOT_ASSIGNED


def test_many_items(make_i):
    items = [(f"架空講座{i}", 1, "回", 1000, "10%", 1000) for i in range(6)]
    record = process_i(make_i(items=items, summary={"net_amount": 6000, "tax_amount": 600,
                                                    "withholding_tax": 612, "document_total": 5988}))
    assert len(record["line_items"]) == 6
    assert codes(record) == NOT_ASSIGNED


def test_withholding_blank_is_unknown(make_i):
    record = process_i(make_i(summary={"withholding_tax": None}))
    assert record["withholding_status"] == "unknown"
    assert record["withholding_tax"] is None


def test_line_sum_mismatch_uses_existing_warning(make_i):
    record = process_i(make_i(summary={"net_amount": 12000}))
    assert ("LINE_ITEMS_SUM_MISMATCH", "net_amount") in codes(record)


def test_without_tax_breakdown_tax_treatment_is_unknown(make_i):
    record = process_i(make_i(include_breakdown=False))
    assert record["tax_treatment"] == "unknown"
    assert ("FIELD_MISSING", "tax_treatment") in codes(record)


def test_serial_and_string_dates(make_i):
    record = process_i(make_i(document_date=to_excel(datetime(2026, 6, 30))))
    assert record["document_date"] == "2026-06-30"
    record = process_i(make_i(document_date="2026年6月30日"))
    assert record["document_date"] == "2026-06-30"
    record = process_i(make_i(document_date="未定"))
    assert ("DATE_INVALID", "document_date") in codes(record)


@pytest.mark.parametrize("label", ["単位", "税率", "金額(円)"])
def test_missing_header_label_is_not_guessed(make_i, label):
    path = make_i(hook=clear_label(label))
    record = process_i(path)
    assert extract_i(path)["template_pattern"] == "instructor_unidentified"
    assert ("TEMPLATE_LABEL_MISMATCH", "line_items") in codes(record)
    assert record["line_items"] is None
    assert record["net_amount"] is None
    assert record["withholding_status"] == "unknown"


@pytest.mark.parametrize("label", ["小計", "消費税", "源泉税額", "合計"])
def test_missing_summary_label_is_not_guessed(make_i, label):
    path = make_i(hook=clear_label(label))
    assert extract_i(path)["template_pattern"] == "instructor_unidentified"
    assert process_i(path)["net_amount"] is None


def test_duplicate_labels_are_not_guessed(make_i):
    assert extract_i(make_i(overrides={"F40": "小計"}))["template_pattern"] == "instructor_unidentified"
    assert extract_i(make_i(overrides={"I20": "数量"}))["template_pattern"] == "instructor_unidentified"


def test_column_structure_mismatch(make_i):
    # ラベルはそろっているが、確認済みの列と違う
    def swap_quantity_and_unit(sheet):
        sheet["D20"], sheet["E20"] = "単位", "数量"
    assert extract_i(make_i(hook=swap_quantity_and_unit))["template_pattern"] == "instructor_unidentified"

    def move_summary(sheet):
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == "小計":
                    cell.value = None
                    sheet[f"E{cell.row}"] = "小計"
    assert extract_i(make_i(hook=move_summary))["template_pattern"] == "instructor_unidentified"


def test_missing_date_label(make_i):
    record = process_i(make_i(hook=clear_label("請求日")))
    assert ("TEMPLATE_LABEL_MISMATCH", "document_date") in codes(record)
    assert record["document_date"] is None


def test_not_confused_with_template_c(make_i):
    path = make_i()
    with pytest.raises(SheetNotFoundError):
        template_c.read(path)
    from blue_return_ai.sample_template_c import write_sample as write_c
    c_path = write_c(path.parent / "c.xlsx", pattern="A")
    with pytest.raises(SheetNotFoundError):
        template_instructor.read(c_path)


@pytest.mark.parametrize(
    ("value", "expected"),
    [("10%", (10, None)), ("８％", (8, None)), (0.1, (10, None)), (10, (10, None)), (None, (None, None)),
     ("10", (None, "FIELD_MISSING")), ("非課税", (None, "FIELD_MISSING")), (0.105, (None, "FIELD_MISSING"))],
)
def test_parse_tax_rate(value, expected):
    assert parse_tax_rate(CellValue(value=value, is_formula=False)) == expected


def test_percent_formatted_tax_rate_cell(make_i):
    record = process_i(make_i(items=[("架空講座", 1, "回", 13000, 0.1, 13000)]))
    assert record["line_items"][0]["tax_rate"] == 10

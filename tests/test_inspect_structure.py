"""構造 Inspector のテスト。対象ファイルはすべてテスト内で作る完全な架空データ。"""

from __future__ import annotations

import json
from datetime import datetime

import pytest
from openpyxl import Workbook

from blue_return_ai.cli import UnsafeOutputDirError
from blue_return_ai.inspect_structure import (
    VALUE_WARNING,
    InspectError,
    inspect_csv,
    inspect_excel,
    main,
    number_format_category,
)
from conftest import REPO_ROOT

# 架空の「個人情報・取引内容」。デフォルト出力に現れてはいけない
FAKE_NAME = "架空花子"
FAKE_ADDRESS = "架空県架空市1-2-3"
FAKE_CLIENT = "架空検査商事株式会社"
FAKE_AMOUNT = 98765
FAKE_DATE = datetime(2026, 7, 15)
FAKE_SECRETS = [FAKE_NAME, FAKE_ADDRESS, FAKE_CLIENT, "98765", "98,765", "2026-07-15", "2026/07/15"]


def build_excel(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "架空シート"
    sheet["A1"] = "御請求書"
    sheet.merge_cells("A1:D1")
    sheet["A3"] = "請求日"
    sheet["B3"] = FAKE_DATE
    sheet["B3"].number_format = "yyyy/mm/dd"
    sheet["A4"] = "宛先"
    sheet["B4"] = FAKE_CLIENT
    sheet["A5"] = "氏名"
    sheet["B5"] = FAKE_NAME
    sheet["B6"] = FAKE_ADDRESS
    sheet["A10"] = "小計"
    sheet["B10"] = FAKE_AMOUNT
    sheet["B10"].number_format = "#,##0"
    sheet["A11"] = "合計"
    sheet["B11"] = "=B10*1"
    sheet["B11"].number_format = "#,##0"
    sheet.column_dimensions["E"].hidden = True
    sheet.row_dimensions[20].hidden = True
    workbook.save(path)
    return path


CSV_LINES = [
    "架空銀行 入出金明細",
    "口座番号,1234567",
    "",
    "取引日,摘要,お支払金額,お預り金額,残高",
    f"2026/07/15,{FAKE_CLIENT},,\"98,765\",\"1,234,567\"",
    f"2026/07/16,カ）{FAKE_NAME},5000,,\"1,229,567\"",
    "2026/07/17,架空カード引落,-12000,,\"1,217,567\"",
]


def write_csv_file(path, encoding="cp932", newline="\r\n", lines=CSV_LINES, bom=False):
    data = newline.join(lines) + newline
    raw = data.encode(encoding)
    if bom:
        raw = b"\xef\xbb\xbf" + raw
    path.write_bytes(raw)
    return path


@pytest.fixture
def excel_file(tmp_path):
    return build_excel(tmp_path / "架空ファイル名_請求書.xlsx")


@pytest.fixture
def csv_file(tmp_path):
    return write_csv_file(tmp_path / "架空ファイル名_銀行.csv")


def run_cli(capsys, args):
    code = main(args)
    captured = capsys.readouterr()
    return code, captured.out, captured.err


# ---------------------------------------------------------------- Excel

def test_excel_default_output_has_structure_but_no_values(excel_file, capsys):
    code, out, err = run_cli(capsys, ["excel", str(excel_file)])
    assert code == 0
    for secret in FAKE_SECRETS + ["請求日", "御請求書", "B10*1", excel_file.name]:
        assert secret not in out + err
    assert "シート数: 1" in out
    assert "A1:D1" in out
    assert "B3 datetime -" in out
    assert "B5 string -" in out
    assert "B10 number -" in out
    assert "B11 formula formula" in out
    assert "非表示列: E" in out


def test_excel_result_contains_no_values(excel_file):
    result = inspect_excel(excel_file)
    text = json.dumps(result, ensure_ascii=False, default=str)
    for secret in FAKE_SECRETS:
        assert secret not in text
    cell = next(c for c in result["sheets"][0]["cells"] if c["cell"] == "B4")
    assert cell == {"cell": "B4", "type": "string", "formula": False}


def test_excel_formula_text_requires_option(excel_file, capsys):
    code, out, _ = run_cli(capsys, ["excel", str(excel_file), "--show-formulas"])
    assert code == 0
    assert "formula_text==B10*1" in out
    assert "警告" in out


def test_excel_with_format_shows_category_only(excel_file, capsys):
    code, out, _ = run_cli(capsys, ["excel", str(excel_file), "--with-format"])
    assert code == 0
    assert "B3 datetime - format=date" in out
    assert "B10 number - format=number" in out
    assert "B11 formula formula format=number cached=empty" in out
    assert "#,##0" not in out


def test_excel_selected_cells_without_show_values_hide_values(excel_file, capsys):
    code, out, _ = run_cli(capsys, ["excel", str(excel_file), "--cells", "a4,B5,B10"])
    assert code == 0
    assert "指定したセル:" in out
    for secret in FAKE_SECRETS + ["宛先"]:
        assert secret not in out
    assert VALUE_WARNING not in out


def test_excel_show_values_only_for_selected_cells_with_warning(excel_file, capsys):
    code, out, err = run_cli(capsys, ["excel", str(excel_file), "--cells", "A3,A10", "--show-values"])
    assert code == 0
    assert "value='請求日'" in out
    assert "value='小計'" in out
    assert VALUE_WARNING in out
    assert VALUE_WARNING in err
    # 指定していないセルの値は出ない
    for secret in FAKE_SECRETS:
        assert secret not in out


def test_excel_show_values_requires_cells(excel_file, capsys):
    code, out, err = run_cli(capsys, ["excel", str(excel_file), "--show-values"])
    assert code == 2
    assert FAKE_NAME not in out + err


def test_excel_find_text_returns_addresses_only(excel_file, capsys):
    code, out, _ = run_cli(capsys, ["excel", str(excel_file), "--find-text", "小計",
                                    "--find-text", "存在しないラベル"])
    assert code == 0
    assert "「小計」: A10" in out
    assert "「存在しないラベル」: 見つかりません" in out
    for secret in FAKE_SECRETS:
        assert secret not in out


def test_excel_invalid_cells_and_sheet_are_rejected(excel_file):
    with pytest.raises(InspectError):
        inspect_excel(excel_file, cells=["A1:B2"])
    with pytest.raises(InspectError):
        inspect_excel(excel_file, sheet="存在しないシート")
    with pytest.raises(InspectError):
        inspect_excel(excel_file, sheet="5")


def test_excel_sha256_is_optional(excel_file, capsys):
    import hashlib
    expected = hashlib.sha256(excel_file.read_bytes()).hexdigest()
    _, out, _ = run_cli(capsys, ["excel", str(excel_file)])
    assert expected not in out
    _, out, _ = run_cli(capsys, ["excel", str(excel_file), "--sha256"])
    assert f"SHA-256: {expected}" in out


def test_excel_source_file_is_not_modified(excel_file, capsys):
    before = excel_file.read_bytes()
    mtime = excel_file.stat().st_mtime_ns
    run_cli(capsys, ["excel", str(excel_file), "--with-format", "--show-formulas",
                     "--cells", "A3", "--show-values"])
    assert excel_file.read_bytes() == before
    assert excel_file.stat().st_mtime_ns == mtime


@pytest.mark.parametrize(
    ("number_format", "expected"),
    [("General", "general"), ("@", "text"), ("yyyy/mm/dd", "date"), ("#,##0", "number"),
     ('"¥"#,##0', "currency"), ("0.0%", "percent"), ("h:mm", "time"),
     ("yyyy/m/d h:mm", "datetime"), ('[$-411]ggge"年"m"月"d"日"', "date"),
     ('#,##0"円（架空）"', "number"), (None, "general")],
)
def test_number_format_category(number_format, expected):
    assert number_format_category(number_format) == expected


# ---------------------------------------------------------------- CSV

def test_csv_default_output_has_structure_but_no_values(csv_file, capsys):
    code, out, err = run_cli(capsys, ["csv", str(csv_file)])
    assert code == 0
    for secret in FAKE_SECRETS + ["1234567", "取引日", "摘要", "架空銀行", csv_file.name]:
        assert secret not in out + err
    assert "文字コード: cp932（推定）" in out
    assert "BOM: なし" in out
    assert "CRLF=7" in out
    assert "区切り文字: comma" in out
    assert "空行: 1" in out
    assert "ヘッダー候補行: [4]" in out
    assert "最頻の列数: 5" in out


def test_csv_result_profiles_columns_without_values(csv_file):
    result = inspect_csv(csv_file)
    text = json.dumps(result, ensure_ascii=False)
    for secret in FAKE_SECRETS + ["1234567", "12000"]:
        assert secret not in text
    assert result["column_count_distribution"] == {1: 1, 2: 1, 5: 4}
    assert result["rows_before_header"] == [{"row": 1, "columns": 1}, {"row": 2, "columns": 2},
                                            {"row": 3, "columns": 0}]
    assert result["data_rows_after_header"] == 3
    date_column = result["columns"][0]
    assert date_column["kinds"] == {"date_like": 3}
    assert date_column["date_shapes"] == {"9999/99/99": 3}
    amount_column = result["columns"][2]
    assert amount_column["kinds"] == {"empty": 1, "integer_like": 2}
    assert amount_column["number_traits"] == {"plain": 1, "sign": 1}
    assert result["columns"][4]["number_traits"] == {"comma": 3}


def test_csv_header_row_requires_option_and_warns(csv_file, capsys):
    code, out, err = run_cli(capsys, ["csv", str(csv_file), "--header-row", "4"])
    assert code == 0
    assert "列 1: 取引日" in out
    assert "列 3: お支払金額" in out
    assert VALUE_WARNING in out
    assert VALUE_WARNING in err
    # ヘッダー以外の行の値は出ない
    for secret in FAKE_SECRETS:
        assert secret not in out


def test_csv_show_row_warns(csv_file, capsys):
    code, out, _ = run_cli(capsys, ["csv", str(csv_file), "--show-row", "5"])
    assert code == 0
    assert FAKE_CLIENT in out
    assert VALUE_WARNING in out


def test_csv_out_of_range_row_is_rejected(csv_file, capsys):
    code, out, err = run_cli(capsys, ["csv", str(csv_file), "--show-row", "99"])
    assert code == 2
    assert FAKE_CLIENT not in out + err


def test_csv_utf8_bom_tab_and_lf(tmp_path):
    lines = ["利用日\t利用先\t利用金額", f"2026/07/01\t{FAKE_CLIENT}\t3,000",
             "2026/07/02\t架空書店\t1,500"]
    path = write_csv_file(tmp_path / "card.csv", encoding="utf-8", newline="\n",
                          lines=lines, bom=True)
    result = inspect_csv(path)
    assert result["encoding"] == "utf-8-sig"
    assert result["bom"] == "utf-8"
    assert result["delimiter"] == "tab"
    assert result["newlines"] == {"CRLF": 0, "LF": 3, "CR": 0}
    assert result["header_candidate_rows"] == [1]


def test_csv_source_file_is_not_modified(csv_file, capsys):
    before = csv_file.read_bytes()
    mtime = csv_file.stat().st_mtime_ns
    run_cli(capsys, ["csv", str(csv_file), "--header-row", "4", "--show-row", "5"])
    assert csv_file.read_bytes() == before
    assert csv_file.stat().st_mtime_ns == mtime


# ---------------------------------------------------------------- 出力ファイル・エラー

def test_output_json_is_written_and_not_overwritten(excel_file, tmp_path, capsys):
    output = tmp_path / "out" / "structure.json"
    code, out, _ = run_cli(capsys, ["excel", str(excel_file), "--output", str(output)])
    assert code == 0
    assert "値は含みません" in out
    data = json.loads(output.read_text(encoding="utf-8"))
    assert data["kind"] == "excel"
    code, _, _ = run_cli(capsys, ["excel", str(excel_file), "--output", str(output)])
    assert code == 1  # 既存ファイルは上書きしない


def test_output_inside_repo_must_be_under_data(excel_file, capsys):
    target = REPO_ROOT / "sample_data" / "should_not_exist.json"
    code, _, err = run_cli(capsys, ["excel", str(excel_file), "--output", str(target)])
    assert code == 2
    assert not target.exists()
    with pytest.raises(UnsafeOutputDirError):
        from blue_return_ai.inspect_structure import _write_json
        _write_json({}, target)


def test_errors_do_not_show_file_name(tmp_path, capsys):
    missing = tmp_path / "存在しない架空ファイル.xlsx"
    code, out, err = run_cli(capsys, ["excel", str(missing)])
    assert code == 1
    assert "FileNotFoundError" in err
    assert missing.name not in out + err
    broken = tmp_path / "壊れた架空ファイル.xlsx"
    broken.write_bytes(b"not an excel file")
    code, out, err = run_cli(capsys, ["excel", str(broken)])
    assert code == 1
    assert broken.name not in out + err

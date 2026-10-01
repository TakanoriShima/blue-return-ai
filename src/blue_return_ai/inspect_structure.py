"""実データの「構造」だけを調べる Inspector（本人がローカルで実行する）。

    python -m blue_return_ai.inspect_structure excel <path> [options]
    python -m blue_return_ai.inspect_structure csv <path> [options]

目的: パーサー実装に必要な構造情報（セル番地・型・数式の有無、CSV の文字コード・列数など）を、
実データの値を表示せずに取得する。

- デフォルトでは、セル・行の値を一切表示しない。ファイル名も表示しない。
- 値を表示するのは、人間が対象（セル・行）と表示オプションを明示した場合だけ。
  その場合は「外部 AI に貼らない」旨の警告を必ず表示する。
- 元ファイルはバイト列として読むだけで、保存・変更しない。
- 結果をファイルに保存する場合、リポジトリ内では data/ 配下にのみ書き込む（上書きしない）。

Claude Code はこのツールを実データ（data/ 配下）に対して実行しない。
"""

from __future__ import annotations

import argparse
import codecs
import csv
import hashlib
import io
import json
import re
import sys
from collections import Counter
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from .cli import UnsafeOutputDirError, check_out_dir

VALUE_WARNING = (
    "!!! 警告: 以下の出力には実データ（個人情報・取引内容）が含まれる可能性があります。"
    "Claude Code 等の外部 AI やチャットに貼り付けないでください。 !!!"
)
FORMULA_WARNING = (
    "!!! 警告: 数式には文字列や金額などの値が直接書かれている場合があります。"
    "共有する前に内容を確認してください。 !!!"
)
SHEET_NAME_NOTE = "※ シート名は表示されます。シート名に取引先名等が含まれていないか、共有前に確認してください。"

_CELL_PATTERN = re.compile(r"^[A-Z]{1,3}[1-9][0-9]*$")


class InspectError(Exception):
    """入力の誤り。メッセージに実データの値・ファイル名を含めない。"""


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ================================================================ Excel

def _value_type(value: Any) -> str:
    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, datetime):
        return "datetime"
    if isinstance(value, date):
        return "date"
    if isinstance(value, time):
        return "time"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "string"
    return "other"


def number_format_category(number_format: str | None) -> str:
    """表示形式の「種類」だけを返す（独自の表示形式に含まれる文字列は返さない）。"""
    if not number_format or number_format == "General":
        return "general"
    if number_format == "@":
        return "text"
    stripped = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", number_format).lower()
    has_time = bool(re.search(r"[hs]", stripped))
    has_date = bool(re.search(r"[yd]", stripped)) or (
        "m" in stripped and not has_time and not re.search(r"[0#?]", stripped))
    if has_date and has_time:
        return "datetime"
    if has_date:
        return "date"
    if has_time:
        return "time"
    if "%" in stripped:
        return "percent"
    if re.search(r"[¥$€£]", number_format) or "[$" in number_format:
        return "currency"
    if re.search(r"[0#?]", stripped):
        return "number"
    return "custom"


def _select_sheets(workbook, sheet: str | None) -> list:
    if sheet is None:
        return list(workbook.worksheets)
    if sheet.isdigit():
        index = int(sheet)
        if not 1 <= index <= len(workbook.worksheets):
            raise InspectError("--sheet の番号が範囲外です")
        return [workbook.worksheets[index - 1]]
    if sheet not in workbook.sheetnames:
        raise InspectError("--sheet で指定したシートが見つかりません")
    return [workbook[sheet]]


def inspect_excel(
    path: Path,
    *,
    sheet: str | None = None,
    cells: list[str] | None = None,
    show_values: bool = False,
    show_formulas: bool = False,
    with_format: bool = False,
    find_texts: list[str] | None = None,
    include_sha256: bool = False,
) -> dict:
    """Excel の構造を辞書で返す。値は show_values かつ cells 指定時のみ含める。"""
    if show_values and not cells:
        raise InspectError("--show-values は --cells と一緒に指定してください（全セルの値は表示しません）")
    for coordinate in cells or []:
        if not _CELL_PATTERN.match(coordinate):
            raise InspectError("--cells は A1 形式のセル番地をカンマ区切りで指定してください")

    data = Path(path).read_bytes()
    workbook = load_workbook(io.BytesIO(data))
    cached = load_workbook(io.BytesIO(data), data_only=True)
    try:
        targets = _select_sheets(workbook, sheet)
        if cells and len(targets) != 1:
            raise InspectError("複数シートのブックで --cells を使う場合は --sheet を指定してください")

        result: dict[str, Any] = {
            "kind": "excel",
            "sheet_count": len(workbook.worksheets),
            "sheets": [],
        }
        if include_sha256:
            result["sha256"] = _sha256(data)

        for target in targets:
            number = workbook.worksheets.index(target) + 1
            cached_sheet = cached.worksheets[number - 1]
            info: dict[str, Any] = {
                "index": number,
                "name": target.title,
                "state": target.sheet_state,
                "max_row": target.max_row,
                "max_column": target.max_column,
                "merged_ranges": sorted(str(r) for r in target.merged_cells.ranges),
                "hidden_rows": sorted(k for k, d in target.row_dimensions.items() if d.hidden),
                "hidden_columns": sorted(k for k, d in target.column_dimensions.items() if d.hidden),
                "cells": [],
            }
            for row in target.iter_rows():
                for cell in row:
                    if isinstance(cell, MergedCell) or cell.value is None:
                        continue
                    info["cells"].append(_cell_info(
                        cell, cached_sheet, show_value=False,
                        show_formula=show_formulas, with_format=with_format))

            if find_texts:
                wanted = {t.strip() for t in find_texts}
                info["found_texts"] = [
                    {"text": text, "cells": [
                        cell.coordinate
                        for row in target.iter_rows() for cell in row
                        if isinstance(cell.value, str) and cell.data_type != "f"
                        and cell.value.strip() == text
                    ]}
                    for text in sorted(wanted)
                ]

            if cells:
                info["selected_cells"] = [
                    _cell_info(target[coordinate], cached_sheet, show_value=show_values,
                               show_formula=show_formulas, with_format=True)
                    for coordinate in cells
                ]
            result["sheets"].append(info)
        return result
    finally:
        workbook.close()
        cached.close()


def _cell_info(cell, cached_sheet, *, show_value: bool, show_formula: bool, with_format: bool) -> dict:
    is_formula = cell.data_type == "f"
    info: dict[str, Any] = {
        "cell": cell.coordinate,
        "type": "formula" if is_formula else _value_type(cell.value),
        "formula": is_formula,
    }
    if isinstance(cell, MergedCell):
        info["merged_non_anchor"] = True
    if with_format:
        info["number_format"] = number_format_category(cell.number_format)
        if is_formula:
            info["cached_type"] = _value_type(cached_sheet[cell.coordinate].value)
    if show_formula and is_formula:
        info["formula_text"] = str(cell.value)
    if show_value:
        value = cached_sheet[cell.coordinate].value if is_formula else cell.value
        info["value"] = value if isinstance(value, (str, int, float, bool)) or value is None else str(value)
    return info


def format_excel_report(result: dict) -> str:
    lines = ["[Excel 構造]（セルの値は表示していません）", f"シート数: {result['sheet_count']}"]
    if "sha256" in result:
        lines.append(f"SHA-256: {result['sha256']}")
    lines.append(SHEET_NAME_NOTE)
    for sheet in result["sheets"]:
        lines.append("")
        lines.append(f"シート {sheet['index']}: name={sheet['name']} state={sheet['state']} "
                     f"max_row={sheet['max_row']} max_column={sheet['max_column']}")
        lines.append(f"  結合セル: {', '.join(sheet['merged_ranges']) or 'なし'}")
        lines.append(f"  非表示行: {', '.join(map(str, sheet['hidden_rows'])) or 'なし'}"
                     f" / 非表示列: {', '.join(sheet['hidden_columns']) or 'なし'}")
        lines.append(f"  非空セル {len(sheet['cells'])} 件（番地 型 数式）:")
        for cell in sheet["cells"]:
            lines.append("    " + _format_cell(cell))
        if "found_texts" in sheet:
            lines.append("  指定した文字列と完全一致するセル番地:")
            for found in sheet["found_texts"]:
                lines.append(f"    「{found['text']}」: {', '.join(found['cells']) or '見つかりません'}")
        if "selected_cells" in sheet:
            lines.append("  指定したセル:")
            for cell in sheet["selected_cells"]:
                lines.append("    " + _format_cell(cell))
    return "\n".join(lines)


def _format_cell(cell: dict) -> str:
    parts = [cell["cell"], cell["type"], "formula" if cell["formula"] else "-"]
    if "number_format" in cell:
        parts.append(f"format={cell['number_format']}")
    if "cached_type" in cell:
        parts.append(f"cached={cell['cached_type']}")
    if cell.get("merged_non_anchor"):
        parts.append("merged(non-anchor)")
    if "formula_text" in cell:
        parts.append(f"formula_text={cell['formula_text']}")
    if "value" in cell:
        parts.append(f"value={cell['value']!r}")
    return " ".join(parts)


# ================================================================ CSV

_DELIMITERS = {",": "comma", "\t": "tab", ";": "semicolon", "|": "pipe"}
_INT_PATTERN = re.compile(r"^[+\-−]?[¥\\]?(\d{1,3}(,\d{3})+|\d+)円?$")
_DECIMAL_PATTERN = re.compile(r"^[+\-−]?[¥\\]?(\d{1,3}(,\d{3})+|\d+)\.\d+円?$")
_DATE_PATTERNS = [
    re.compile(r"^\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2}$"),
    re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日$"),
    re.compile(r"^\d{8}$"),
    re.compile(r"^\d{2}[/\-.]\d{1,2}[/\-.]\d{1,2}$"),
    re.compile(r"^\d{1,2}[/\-]\d{1,2}$"),
    re.compile(r"^[RHＲＨ令平]\D{0,2}\d{1,2}[/\-.年]\d{1,2}[/\-.月]\d{1,2}日?$"),
    re.compile(r"^\d{4}[/\-.]\d{1,2}$"),
    re.compile(r"^\d{4}年\d{1,2}月$"),
]


def _field_kind(text: str) -> str:
    value = text.strip().strip('"')
    if not value:
        return "empty"
    if any(p.match(value) for p in _DATE_PATTERNS):
        return "date_like"
    if _INT_PATTERN.match(value):
        return "integer_like"
    if _DECIMAL_PATTERN.match(value):
        return "decimal_like"
    return "text"


def _date_shape(text: str) -> str:
    """日付らしい値の書式だけを返す（数字は 9 に置き換える）。"""
    return re.sub(r"\d", "9", text.strip().strip('"'))


def _number_traits(text: str) -> tuple[str, ...]:
    value = text.strip().strip('"')
    traits = []
    if "," in value:
        traits.append("comma")
    if value[:1] in "+-−":
        traits.append("sign")
    if "." in value:
        traits.append("decimal")
    if "¥" in value or "\\" in value or "円" in value:
        traits.append("yen")
    return tuple(traits) or ("plain",)


def _detect_encoding(data: bytes, forced: str | None) -> tuple[str, str | None, bool]:
    """(デコードに使う文字コード, BOM の種類, 推定かどうか) を返す。"""
    if data.startswith(codecs.BOM_UTF8):
        return "utf-8-sig", "utf-8", False
    if data.startswith(codecs.BOM_UTF16_LE) or data.startswith(codecs.BOM_UTF16_BE):
        return "utf-16", "utf-16", False
    if forced:
        return forced, None, False
    for candidate in ("utf-8", "cp932"):
        try:
            data.decode(candidate)
            return candidate, None, True
        except UnicodeDecodeError:
            continue
    return "unknown", None, True


def _newline_style(text: str) -> dict[str, int]:
    crlf = text.count("\r\n")
    return {"CRLF": crlf, "LF": text.count("\n") - crlf, "CR": text.count("\r") - crlf}


def _parse_rows(text: str, delimiter: str) -> list[list[str]]:
    return list(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter))


def _choose_delimiter(text: str) -> tuple[str, dict[str, int]]:
    """各候補で「列数が最頻値の行」の数を数え、最も多い区切り文字を選ぶ。"""
    scores: dict[str, int] = {}
    for delimiter in _DELIMITERS:
        counts = Counter(len(r) for r in _parse_rows(text, delimiter) if any(f.strip() for f in r))
        if not counts:
            scores[delimiter] = 0
            continue
        modal_columns, modal_rows = counts.most_common(1)[0]
        scores[delimiter] = modal_rows if modal_columns > 1 else 0
    best = max(scores, key=lambda d: (scores[d], d == ","))
    return best, {_DELIMITERS[d]: s for d, s in scores.items()}


def inspect_csv(
    path: Path,
    *,
    encoding: str | None = None,
    delimiter: str | None = None,
    show_rows: list[int] | None = None,
    header_row: int | None = None,
    include_sha256: bool = False,
) -> dict:
    """CSV の構造を辞書で返す。行の値は show_rows / header_row 指定時のみ含める。"""
    data = Path(path).read_bytes()
    used_encoding, bom, estimated = _detect_encoding(data, encoding)
    decode_errors = 0
    if used_encoding == "unknown":
        text = data.decode("utf-8", errors="replace")
        decode_errors = text.count("�")
    else:
        try:
            text = data.decode(used_encoding)
        except (UnicodeDecodeError, LookupError):
            raise InspectError("指定した文字コードでデコードできません") from None

    chosen, scores = (delimiter, {}) if delimiter else _choose_delimiter(text)
    rows = _parse_rows(text, chosen)
    non_empty = [r for r in rows if any(f.strip() for f in r)]
    distribution = Counter(len(r) for r in non_empty)
    modal_columns = distribution.most_common(1)[0][0] if distribution else 0

    header_candidates = []
    for number, row in enumerate(rows, start=1):
        kinds = [_field_kind(f) for f in row]
        filled = [k for k in kinds if k != "empty"]
        if len(row) == modal_columns and filled and all(k == "text" for k in filled):
            header_candidates.append(number)
        if len(header_candidates) >= 3:
            break

    data_start = header_candidates[0] + 1 if header_candidates else 1
    columns: list[dict[str, Any]] = [
        {"column": i + 1, "kinds": Counter(), "date_shapes": Counter(), "number_traits": Counter()}
        for i in range(modal_columns)
    ]
    for row in rows[data_start - 1:]:
        if len(row) != modal_columns:
            continue
        for i, field in enumerate(row):
            kind = _field_kind(field)
            columns[i]["kinds"][kind] += 1
            if kind == "date_like":
                columns[i]["date_shapes"][_date_shape(field)] += 1
            elif kind in ("integer_like", "decimal_like"):
                columns[i]["number_traits"]["+".join(_number_traits(field))] += 1

    result: dict[str, Any] = {
        "kind": "csv",
        "encoding": used_encoding,
        "encoding_estimated": estimated,
        "bom": bom,
        "decode_errors": decode_errors,
        "newlines": _newline_style(text),
        "delimiter": _DELIMITERS.get(chosen, "custom"),
        "delimiter_scores": scores,
        "physical_lines": len(text.splitlines()),
        "rows": len(rows),
        "empty_rows": len(rows) - len(non_empty),
        "column_count_distribution": dict(sorted(distribution.items())),
        "modal_column_count": modal_columns,
        "header_candidate_rows": header_candidates,
        "rows_before_header": [
            {"row": n, "columns": len(rows[n - 1])}
            for n in range(1, (header_candidates[0] if header_candidates else 1))
        ],
        "data_rows_after_header": sum(
            1 for r in rows[data_start - 1:] if any(f.strip() for f in r)),
        "columns": [
            {
                "column": c["column"],
                "kinds": dict(c["kinds"]),
                "date_shapes": dict(c["date_shapes"].most_common(3)),
                "number_traits": dict(c["number_traits"]),
            }
            for c in columns
        ],
    }
    if include_sha256:
        result["sha256"] = _sha256(data)

    requested = sorted(set((show_rows or []) + ([header_row] if header_row else [])))
    for number in requested:
        if not 1 <= number <= len(rows):
            raise InspectError("指定した行番号が範囲外です")
    if show_rows:
        result["shown_rows"] = [{"row": n, "fields": rows[n - 1]} for n in sorted(set(show_rows))]
    if header_row:
        result["header"] = {"row": header_row, "columns": rows[header_row - 1]}
    return result


def format_csv_report(result: dict) -> str:
    lines = ["[CSV 構造]（行の値は表示していません）"]
    if "sha256" in result:
        lines.append(f"SHA-256: {result['sha256']}")
    lines += [
        f"文字コード: {result['encoding']}{'（推定）' if result['encoding_estimated'] else ''}"
        f" / BOM: {result['bom'] or 'なし'} / デコードできない文字: {result['decode_errors']}",
        f"改行: {', '.join(f'{k}={v}' for k, v in result['newlines'].items())}",
        f"区切り文字: {result['delimiter']}"
        + (f"（候補ごとの一致行数: {result['delimiter_scores']}）" if result["delimiter_scores"] else ""),
        f"物理行数: {result['physical_lines']} / CSV 行数: {result['rows']} / 空行: {result['empty_rows']}",
        f"列数の分布（列数: 行数）: {result['column_count_distribution']}",
        f"最頻の列数: {result['modal_column_count']}",
        f"ヘッダー候補行: {result['header_candidate_rows'] or 'なし'}",
        f"ヘッダー候補より前の行（行番号: 列数）: "
        f"{', '.join(f'{r['row']}:{r['columns']}' for r in result['rows_before_header']) or 'なし'}",
        f"ヘッダー候補より後の非空行数: {result['data_rows_after_header']}",
        "列ごとの値の種類（値そのものは表示しません）:",
    ]
    for column in result["columns"]:
        line = f"  列 {column['column']}: {column['kinds']}"
        if column["date_shapes"]:
            line += f" 日付の書式={column['date_shapes']}"
        if column["number_traits"]:
            line += f" 数値の特徴={column['number_traits']}"
        lines.append(line)
    if "header" in result:
        lines.append(f"ヘッダーとして表示（{result['header']['row']} 行目）:")
        for i, name in enumerate(result["header"]["columns"], start=1):
            lines.append(f"  列 {i}: {name}")
    for shown in result.get("shown_rows", []):
        lines.append(f"{shown['row']} 行目の値:")
        for i, value in enumerate(shown["fields"], start=1):
            lines.append(f"  列 {i}: {value}")
    return "\n".join(lines)


# ================================================================ CLI

def _write_json(result: dict, output: Path) -> Path:
    check_out_dir(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2, default=str)
        f.write("\n")
    return output


def _split_cells(text: str | None) -> list[str] | None:
    if not text:
        return None
    return [c.strip().upper() for c in text.split(",") if c.strip()]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m blue_return_ai.inspect_structure",
        description="実データの値を表示せずに、Excel / CSV の構造だけを調べる（本人がローカルで実行する）")
    sub = parser.add_subparsers(dest="kind", required=True)

    excel = sub.add_parser("excel", help="Excel（.xlsx）の構造")
    excel.add_argument("path", type=Path)
    excel.add_argument("--sheet", help="対象シート（シート名または 1 始まりの番号）")
    excel.add_argument("--cells", help="詳細を見るセル番地（例: L11,M11,A20）。値は --show-values 指定時のみ")
    excel.add_argument("--show-values", action="store_true",
                       help="--cells で指定したセルの値を表示する（実データが表示されます）")
    excel.add_argument("--show-formulas", action="store_true", help="数式の中身を表示する（値を含む場合があります）")
    excel.add_argument("--with-format", action="store_true",
                       help="表示形式の種類と、数式セルのキャッシュ値の型も表示する")
    excel.add_argument("--find-text", action="append", metavar="TEXT",
                       help="指定した文字列と完全一致するセル番地を表示する（例: --find-text 小計）。複数指定可")

    table = sub.add_parser("csv", help="CSV の構造")
    table.add_argument("path", type=Path)
    table.add_argument("--encoding", help="文字コードを指定する（例: cp932, utf-8）")
    table.add_argument("--delimiter", choices=["comma", "tab", "semicolon", "pipe"])
    table.add_argument("--header-row", type=int, metavar="N",
                       help="N 行目をヘッダーとして列番号付きで表示する（値が表示されます）")
    table.add_argument("--show-row", type=int, action="append", metavar="N",
                       help="N 行目の値を表示する（実データが表示されます）。複数指定可")

    for sub_parser in (excel, table):
        sub_parser.add_argument("--sha256", action="store_true", help="ファイルの SHA-256 を表示する")
        sub_parser.add_argument("--output", type=Path,
                                help="結果を JSON で保存する（リポジトリ内では data/ 配下のみ。上書きしない）")

    args = parser.parse_args(argv)
    try:
        if args.kind == "excel":
            shows_data = args.show_values
            result = inspect_excel(
                args.path, sheet=args.sheet, cells=_split_cells(args.cells),
                show_values=args.show_values, show_formulas=args.show_formulas,
                with_format=args.with_format, find_texts=args.find_text,
                include_sha256=args.sha256)
            report = format_excel_report(result)
        else:
            delimiter = {v: k for k, v in _DELIMITERS.items()}.get(args.delimiter)
            shows_data = bool(args.show_row or args.header_row)
            result = inspect_csv(
                args.path, encoding=args.encoding, delimiter=delimiter,
                show_rows=args.show_row, header_row=args.header_row, include_sha256=args.sha256)
            report = format_csv_report(result)

        warnings = []
        if shows_data:
            warnings.append(VALUE_WARNING)
        if args.kind == "excel" and args.show_formulas:
            warnings.append(FORMULA_WARNING)
        for warning in warnings:
            print(warning, file=sys.stderr)
            print(warning)
        if args.output:
            _write_json(result, args.output)
            note = "実データを含む可能性があります" if shows_data else "セル・行の値は含みません"
            print(f"結果を JSON で保存しました（{note}）")
        else:
            print(report)
        for warning in warnings:
            print(warning)
        return 0
    except (InspectError, UnsafeOutputDirError) as error:
        print(f"エラー: {error}", file=sys.stderr)
        return 2
    except Exception as error:  # noqa: BLE001 - 実データを含む可能性があるため詳細は表示しない
        print(f"エラー: 読み取りに失敗しました: {type(error).__name__}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

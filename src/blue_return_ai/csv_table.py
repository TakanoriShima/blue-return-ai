"""銀行・カード CSV の共通読み込み（docs/data_model.md §17）。

- 書式（文字コード・ヘッダー行・列名）は CsvFormat で指定する。公開コードには汎用の識別子だけを使う。
- ヘッダーより前の行（口座情報・利用者情報等が入り得る）は読み飛ばし、内容を解析しない。
- ヘッダーが書式と一致しない場合は、資料全体を取り込まない（CsvFormatError）。
- 例外・警告に実データの値を含めない。
"""

from __future__ import annotations

import csv
import hashlib
import io
import re
from dataclasses import dataclass, replace
from pathlib import Path

from .cell_parsing import normalize_label, parse_amount, parse_date
from .excel_reader import CellValue

_COMPACT_DATE = re.compile(r"^(\d{4})(\d{2})(\d{2})$")


class CsvFormatError(Exception):
    """ヘッダーの不一致・文字コード違い等。メッセージに実データの値を含めない。"""


@dataclass(frozen=True)
class CsvFormat:
    format_id: str
    encoding: str
    header_row: int
    columns: tuple[str, ...]

    def with_overrides(self, overrides: dict) -> "CsvFormat":
        """設定ファイルからの上書き（encoding / header_row / columns）。"""
        allowed = {"encoding", "header_row", "columns"}
        unknown = set(overrides) - allowed
        if unknown:
            raise ValueError(f"書式で上書きできるのは {', '.join(sorted(allowed))} のみです")
        values = dict(overrides)
        if "encoding" in values and (not isinstance(values["encoding"], str) or not values["encoding"]):
            raise ValueError("encoding は文字列で指定してください")
        if "header_row" in values and (isinstance(values["header_row"], bool)
                                       or not isinstance(values["header_row"], int)
                                       or values["header_row"] < 1):
            raise ValueError("header_row は 1 以上の整数で指定してください")
        if "columns" in values:
            columns = values["columns"]
            if (not isinstance(columns, list) or len(columns) != len(self.columns)
                    or not all(isinstance(c, str) and c for c in columns)):
                raise ValueError("columns は既定と同じ数の列名（文字列）の配列で指定してください")
            values["columns"] = tuple(columns)
        return replace(self, **values)


@dataclass(frozen=True)
class CsvRow:
    row_number: int  # CSV 上の行番号（1 始まり。ヘッダー行を含めて数える）
    fields: list[str]


def read_table(path: Path, fmt: CsvFormat) -> tuple[str, list[CsvRow]]:
    """(source_hash, ヘッダーより後の行) を返す。ヘッダーより前の行は解析しない。"""
    data = Path(path).read_bytes()
    source_hash = hashlib.sha256(data).hexdigest()
    try:
        text = data.decode(fmt.encoding)
    except (UnicodeDecodeError, LookupError):
        raise CsvFormatError("設定した文字コードで読み取れません") from None
    if text.startswith("﻿"):
        text = text[1:]

    rows = list(csv.reader(io.StringIO(text, newline="")))
    if len(rows) < fmt.header_row:
        raise CsvFormatError("ヘッダー行がありません")
    header = [normalize_label(c) for c in rows[fmt.header_row - 1]]
    while header and header[-1] == "":
        header.pop()
    if header != [normalize_label(c) for c in fmt.columns]:
        raise CsvFormatError("ヘッダーが書式と一致しません")

    return source_hash, [
        CsvRow(row_number=number, fields=row)
        for number, row in enumerate(rows[fmt.header_row:], start=fmt.header_row + 1)
    ]


def fit_columns(fields: list[str], count: int) -> list[str] | None:
    """列数を書式に合わせる。末尾の空欄の余分な列は除き、それ以外で列数が違えば None。"""
    fields = list(fields)
    while len(fields) > count and fields[-1].strip() == "":
        fields.pop()
    return fields if len(fields) == count else None


def is_blank(fields: list[str]) -> bool:
    return all(f.strip() == "" for f in fields)


def text_or_none(value: str) -> str | None:
    value = value.strip()
    return value or None


def parse_csv_amount(value: str):
    return parse_amount(CellValue(value=value.strip() or None, is_formula=False))


def parse_csv_date(value: str):
    """CSV の日付文字列を YYYY-MM-DD にする。8 桁（YYYYMMDD）にも対応する。"""
    text = value.strip()
    match = _COMPACT_DATE.match(text)
    if match:
        text = "/".join(match.groups())
    return parse_date(CellValue(value=text or None, is_formula=False))

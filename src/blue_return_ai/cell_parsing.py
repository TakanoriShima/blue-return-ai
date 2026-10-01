"""セルの値を sales_record の型へ変換する。

各関数は (値, 警告コード) を返す。解釈できない値は推測せず None とし、警告コードを返す。
"""

from __future__ import annotations

import math
import re
import unicodedata
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any

from openpyxl.utils.datetime import WINDOWS_EPOCH, from_excel

from .excel_reader import CellValue

Parsed = tuple[Any, str | None]

_REIWA_START_YEAR = 2018  # 令和1年 = 2019年

_DATE_PATTERN = re.compile(r"^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?$")
_REIWA_PATTERN = re.compile(r"^(?:令和|R)(\d{1,2}|元)[.年](\d{1,2})[.月](\d{1,2})日?$")
_NO_YEAR_DATE_PATTERN = re.compile(r"^(\d{1,2})[/.月](\d{1,2})日?$")
_MONTH_PATTERN = re.compile(r"^(\d{4})[-/.年](\d{1,2})月?分?$")
_NO_YEAR_MONTH_PATTERN = re.compile(r"^(\d{1,2})月分?$")
_AMOUNT_PATTERN = re.compile(r"^-?\d+$")
_MAX_EXCEL_SERIAL = 2958465  # Excel で扱える最後の日付（9999-12-31）のシリアル値


def _normalize_text(value: str) -> str:
    return unicodedata.normalize("NFKC", value).strip()


def normalize_label(text: str) -> str:
    """ラベル・見出しの比較用。NFKC 正規化し、空白をすべて除く（完全一致の比較に使う）。"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text))


def _formula_check(cell: CellValue) -> str | None:
    return "FORMULA_NO_CACHED_VALUE" if cell.missing_cached_value else None


def parse_text(cell: CellValue) -> Parsed:
    if (code := _formula_check(cell)) is not None:
        return None, code
    if cell.value is None:
        return None, None
    text = str(cell.value).strip()
    return (text or None), None


def parse_amount(cell: CellValue) -> Parsed:
    """円単位の整数へ変換する。小数・解釈不能な値は丸めずに None とする。"""
    if (code := _formula_check(cell)) is not None:
        return None, code
    value = cell.value
    if value is None:
        return None, None
    if isinstance(value, bool):
        return None, "AMOUNT_NOT_INTEGER"
    if isinstance(value, int):
        return value, None
    if isinstance(value, float):
        return (int(value), None) if value.is_integer() else (None, "AMOUNT_NOT_INTEGER")
    if isinstance(value, str):
        text = _normalize_text(value)
        for symbol in ("¥", "円", ",", " "):
            text = text.replace(symbol, "")
        if not text:
            return None, None
        if _AMOUNT_PATTERN.match(text):
            return int(text), None
    return None, "AMOUNT_NOT_INTEGER"


def parse_quantity(cell: CellValue) -> Parsed:
    """数量を十進数の文字列へ変換する（float のまま保持しない）。"""
    if (code := _formula_check(cell)) is not None:
        return None, code
    value = cell.value
    if value is None:
        return None, None
    if isinstance(value, bool):
        return None, "QUANTITY_INVALID"
    if isinstance(value, int):
        return str(value), None
    if isinstance(value, float):
        return (str(int(value)) if value.is_integer() else repr(value)), None
    if isinstance(value, str):
        text = _normalize_text(value).replace(",", "")
        if not text:
            return None, None
        try:
            number = Decimal(text)
        except InvalidOperation:
            return None, "QUANTITY_INVALID"
        if number.is_finite():
            return text, None
    return None, "QUANTITY_INVALID"


def parse_date(cell: CellValue) -> Parsed:
    """YYYY-MM-DD 形式の文字列へ変換する。年が無い表記は補完しない。"""
    if (code := _formula_check(cell)) is not None:
        return None, code
    value = cell.value
    if value is None:
        return None, None
    if isinstance(value, datetime):
        return value.date().isoformat(), None
    if isinstance(value, date):
        return value.isoformat(), None
    if not isinstance(value, str):
        return None, "DATE_INVALID"

    text = _normalize_text(value)
    if not text:
        return None, None
    if match := _DATE_PATTERN.match(text):
        year, month, day = (int(g) for g in match.groups())
    elif match := _REIWA_PATTERN.match(text):
        era_year = 1 if match.group(1) == "元" else int(match.group(1))
        year = _REIWA_START_YEAR + era_year
        month, day = int(match.group(2)), int(match.group(3))
    elif _NO_YEAR_DATE_PATTERN.match(text):
        return None, "DATE_AMBIGUOUS"
    else:
        return None, "DATE_INVALID"

    try:
        return date(year, month, day).isoformat(), None
    except ValueError:
        return None, "DATE_INVALID"


def parse_month(cell: CellValue) -> Parsed:
    """業務対象月を YYYY-MM 形式の文字列へ変換する。月初・月末の日付へは展開しない。"""
    if (code := _formula_check(cell)) is not None:
        return None, code
    value = cell.value
    if value is None:
        return None, None
    if isinstance(value, (datetime, date)):
        # テンプレート上「月」として定義されたセルのため、年月のみを使う
        return f"{value.year:04d}-{value.month:02d}", None
    if not isinstance(value, str):
        return None, "DATE_INVALID"

    text = _normalize_text(value)
    if not text:
        return None, None
    if match := _MONTH_PATTERN.match(text):
        year, month = int(match.group(1)), int(match.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}", None
        return None, "DATE_INVALID"
    if _NO_YEAR_MONTH_PATTERN.match(text):
        return None, "DATE_AMBIGUOUS"
    return None, "DATE_INVALID"


def parse_excel_date(cell: CellValue, epoch: datetime = WINDOWS_EPOCH) -> Parsed:
    """テンプレート定義で「日付」と決まっているセル用。

    表示形式が日付でない（標準の）数値セルは、Excel の日付シリアル値として openpyxl の
    from_excel（ブックの epoch に従う）で日付に変換する。それ以外（datetime・文字列の日付）は parse_date と同じ。
    範囲外・不正な数値は DATE_INVALID とする。
    """
    if (code := _formula_check(cell)) is not None:
        return None, code
    value = cell.value
    if isinstance(value, bool):
        return None, "DATE_INVALID"
    if isinstance(value, (int, float)):
        if not math.isfinite(value) or not 1 <= value <= _MAX_EXCEL_SERIAL:
            return None, "DATE_INVALID"
        try:
            converted = from_excel(value, epoch=epoch)
        except (ValueError, OverflowError, TypeError):
            return None, "DATE_INVALID"
        if not isinstance(converted, datetime):
            return None, "DATE_INVALID"
        return converted.date().isoformat(), None
    return parse_date(cell)

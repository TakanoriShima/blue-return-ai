"""ラベルを起点に Excel の値を探す補助関数（template_c・template_instructor で共有）。

- ラベルの比較は normalize_label（NFKC 正規化＋空白除去）後の完全一致。
- 値は「ラベルと同じ行で、ラベルより右にある最も近い空でないセル」から取る。
"""

from __future__ import annotations

from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

from .cell_parsing import normalize_label
from .excel_reader import ExcelSource


def split(coordinate: str) -> tuple[str, int]:
    """セル番地を (列の文字, 行番号) に分ける。"""
    column, row = coordinate_from_string(coordinate)
    return column, row


def row_of(coordinate: str) -> int:
    return split(coordinate)[1]


def column_of(coordinate: str) -> str:
    return split(coordinate)[0]


def column_index(coordinate: str) -> int:
    return column_index_from_string(column_of(coordinate))


def _sort_key(coordinate: str) -> tuple[int, int]:
    return row_of(coordinate), column_index(coordinate)


def label_cells(source: ExcelSource) -> dict[str, list[str]]:
    """正規化したラベル → セル番地の一覧（数式でない文字列セルのみ。上から順）。"""
    labels: dict[str, list[str]] = {}
    for coordinate, cell in source.cells.items():
        if isinstance(cell.value, str) and not cell.is_formula:
            labels.setdefault(normalize_label(cell.value), []).append(coordinate)
    for coordinates in labels.values():
        coordinates.sort(key=_sort_key)
    return labels


def labels_in_row(labels: dict[str, list[str]], row: int) -> list[tuple[str, str]]:
    """指定行にある (正規化ラベル, セル番地) の一覧。"""
    return [(label, cell) for label, cells in labels.items() for cell in cells if row_of(cell) == row]


def value_right_of(source: ExcelSource, label_cell: str) -> str | None:
    """ラベルと同じ行で、ラベルより右にある最も近い空でないセル（無ければ None）。

    同じ行のさらに右にある別の欄（例：氏名欄）は読まない。最も近いセルが金額・日付として解釈できなければ、
    各項目の型の警告（AMOUNT_NOT_INTEGER / DATE_INVALID）で検出される。
    """
    row = row_of(label_cell)
    start = column_index(label_cell)
    found = sorted(
        (coordinate for coordinate, cell in source.cells.items()
         if row_of(coordinate) == row and column_index(coordinate) > start
         and (cell.value is not None or cell.is_formula)),
        key=column_index,
    )
    return found[0] if found else None

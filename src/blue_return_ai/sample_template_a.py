"""template_a 形式の「完全な架空」Excel 請求書を生成する（開発・テスト用）。

    python -m blue_return_ai.sample_template_a sample_data/invoices/template_a_sample.xlsx

- 取引先名・件名・金額・日付はすべて架空の値であり、実在の企業・人物・請求書を模倣していない。
- 金額は数式ではなく値として書き込む（openpyxl で作成した数式セルはキャッシュ値を持たないため）。
- 既存ファイルは上書きしない。data/ 配下へは書き込まない。
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from . import template_a

REPO_ROOT = Path(__file__).resolve().parents[2]

# 完全な架空データ
SAMPLE_VALUES: dict[str, Any] = {
    "A1": "請求書",
    "A3": "請求日", "B3": datetime(2026, 2, 28),
    "A4": "業務対象月", "B4": "2026年2月分",
    "A5": "取引先", "B5": "架空サンプル商事株式会社",
    "A6": "件名", "B6": "研修講師業務委託費（2026年2月分）",
    "A7": "支払期限", "B7": datetime(2026, 3, 31),
    "A9": "明細", "B9": "数量", "C9": "単位", "D9": "単価", "E9": "金額",
    "A10": "研修講師業務", "B10": 2, "C10": "日", "D10": 25000, "E10": 50000,
    "A11": "研修教材作成", "B11": 1.5, "C11": "時間", "D11": 4000, "E11": 6000,
    "D16": "税抜金額", "E16": 56000,
    "D17": "消費税額", "E17": 5600,
    "D18": "税込総額", "E18": 61600,
    "D19": "源泉徴収税額", "E19": 5717,
    "D20": "振込予定額", "E20": 55883,
    "A22": "※ この請求書は blue-return-ai の開発・テスト用に作成した完全な架空データです。",
}

DATE_CELLS = ("B3", "B7")
AMOUNT_CELLS = ("D10", "D11", "E10", "E11", "E16", "E17", "E18", "E19", "E20")


def build_workbook(overrides: dict[str, Any] | None = None) -> Workbook:
    """架空の template_a 請求書を作る。overrides でセルの値を差し替えられる（None で空欄）。"""
    values = {**SAMPLE_VALUES, **(overrides or {})}
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = template_a.SHEET_NAME
    for coordinate, value in values.items():
        if value is not None:
            sheet[coordinate] = value
    for coordinate in DATE_CELLS:
        sheet[coordinate].number_format = "yyyy/mm/dd"
    for coordinate in AMOUNT_CELLS:
        sheet[coordinate].number_format = "#,##0"
    workbook.properties.creator = "blue-return-ai sample generator"
    return workbook


def write_sample(path: Path, overrides: dict[str, Any] | None = None) -> Path:
    path = Path(path)
    resolved = path.resolve()
    if resolved.is_relative_to(REPO_ROOT / "data"):
        raise ValueError("架空データを data/ 配下へ書き込むことはできません")
    if path.exists():
        raise FileExistsError("既存ファイルは上書きしません")
    path.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(overrides).save(path)
    return path


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if len(args) != 1:
        print("使い方: python -m blue_return_ai.sample_template_a <出力先.xlsx>", file=sys.stderr)
        return 2
    write_sample(Path(args[0]))
    print(f"架空の template_a 請求書を作成しました: {args[0]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

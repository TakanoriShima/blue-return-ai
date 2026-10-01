"""template_b 形式の「完全な架空」Excel 請求書を生成する（開発・テスト用）。

    python -m blue_return_ai.sample_template_b sample_data/invoices/template_b_sample.xlsx

- 取引先名・件名・金額・日付はすべて架空の値であり、実在の企業・人物・請求書を模倣していない。
- 内税・源泉徴収型の書式。帳票上の「小計」は税込総額を意味し、税抜金額の欄は無い。
- 金額は数式ではなく値として書き込む（openpyxl で作成した数式セルはキャッシュ値を持たないため）。
- 既存ファイルは上書きしない。data/ 配下へは書き込まない。
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from . import template_b

REPO_ROOT = Path(__file__).resolve().parents[2]

# 完全な架空データ
SAMPLE_VALUES: dict[str, Any] = {
    "A1": "御請求書",
    "A3": "宛先", "B3": "架空研修サービス株式会社",
    "A4": "発行日", "B4": datetime(2026, 4, 30),
    "A5": "対象期間", "B5": datetime(2026, 4, 1),
    "A6": "件名", "B6": "研修講師業務（2026年4月分）",
    "A7": "お支払期限", "B7": datetime(2026, 5, 31),
    "B10": "品目", "C10": "数量", "D10": "単位", "E10": "単価", "F10": "金額",
    "B11": "研修講師業務", "C11": 4, "D11": "日", "E11": 3000, "F11": 12000,
    "E17": "小計", "F17": 12000,
    "E18": "（内消費税）", "F18": 1090,
    "E19": "源泉徴収税", "F19": 1113,
    "E20": "お振込金額", "F20": 10887,
    "A22": "※ この請求書は blue-return-ai の開発・テスト用に作成した完全な架空データです。",
}

DATE_CELLS = ("B4", "B7")
MONTH_CELLS = ("B5",)
AMOUNT_CELLS = ("E11", "F11", "F17", "F18", "F19", "F20")


def build_workbook(overrides: dict[str, Any] | None = None) -> Workbook:
    """架空の template_b 請求書を作る。overrides でセルの値を差し替えられる（None で空欄）。"""
    values = {**SAMPLE_VALUES, **(overrides or {})}
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = template_b.SHEET_NAME
    for coordinate, value in values.items():
        if value is not None:
            sheet[coordinate] = value
    for coordinate in DATE_CELLS:
        sheet[coordinate].number_format = "yyyy/mm/dd"
    for coordinate in MONTH_CELLS:
        sheet[coordinate].number_format = 'yyyy"年"m"月分"'
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
        print("使い方: python -m blue_return_ai.sample_template_b <出力先.xlsx>", file=sys.stderr)
        return 2
    write_sample(Path(args[0]))
    print(f"架空の template_b 請求書を作成しました: {args[0]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

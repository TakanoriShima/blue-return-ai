"""sales_record の JSON 出力と確認用 CSV 出力（docs/data_model.md §8.4、§8.5）。

- 出力ファイル名に元資料のファイル名を使わない（record_id・実行日時のみ）。
- 既存ファイルは上書きしない。
"""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any, Iterable

CSV_ENCODING = "utf-8-sig"  # Excel で開きやすい BOM 付き UTF-8
_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

RECORD_COLUMNS = [
    "record_id",
    "review_status",
    "warnings",
    "document_type",
    "source_name",
    "document_date",
    "service_period_from",
    "service_period_to",
    "service_period_month",
    "customer",
    "description",
    "line_item_count",
    "net_amount",
    "tax_amount",
    "gross_amount",
    "tax_treatment",
    "tax_rate",
    "withholding_status",
    "withholding_tax",
    "deductions_total",
    "expected_payment_amount",
    "calculated_fields",
    "payment_due",
    "source_hash",
    # 人間が記入する列（プログラムは常に空欄で出力する）
    "human_review_status",
    "human_revenue_date",
    "human_reviewer_note",
]

LINE_ITEM_COLUMNS = [
    "record_id",
    "line_no",
    "description",
    "quantity",
    "unit",
    "unit_price",
    "amount",
]


def write_record_json(record: dict, directory: Path) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{record['record_id']}.json"
    with path.open("x", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return path


def load_record_json(path: Path) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def find_records_by_source_hash(directory: Path, source_hash: str) -> list[str]:
    """出力済み JSON から、同じ source_hash を持つ record_id を探す（既存取込候補の検出）。"""
    if not directory.is_dir():
        return []
    found = []
    for path in sorted(directory.glob("*.json")):
        record = load_record_json(path)
        if record.get("source_hash") == source_hash:
            found.append(record["record_id"])
    return found


def escape_csv_value(value: Any) -> Any:
    """Excel で数式として解釈されないよう、文字列の先頭に ' を付ける。"""
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIXES):
        return "'" + value
    return value


def _cell(value: Any) -> Any:
    # null は空欄、0 は 0 として出力する
    return "" if value is None else escape_csv_value(value)


def record_to_row(record: dict) -> dict:
    period = record["service_period"] or {}
    deductions = record["deductions"]
    return {
        "record_id": record["record_id"],
        "review_status": record["review_status"],
        "warnings": ";".join(f"{w['code']}:{w['field']}" for w in record["warnings"]),
        "document_type": record["document_type"],
        "source_name": record["source_name"],
        "document_date": record["document_date"],
        "service_period_from": period.get("from"),
        "service_period_to": period.get("to"),
        "service_period_month": period.get("month"),
        "customer": record["customer"],
        "description": record["description"],
        "line_item_count": None if record["line_items"] is None else len(record["line_items"]),
        "net_amount": record["net_amount"],
        "tax_amount": record["tax_amount"],
        "gross_amount": record["gross_amount"],
        "tax_treatment": record["tax_treatment"],
        "tax_rate": record["tax_rate"],
        "withholding_status": record["withholding_status"],
        "withholding_tax": record["withholding_tax"],
        "deductions_total": (
            None if deductions is None
            or any(d.get("amount") is None for d in deductions)
            else sum(d["amount"] for d in deductions)
        ),
        "expected_payment_amount": record["expected_payment_amount"],
        "calculated_fields": ";".join(record["calculated_fields"]),
        "payment_due": record["payment_due"],
        "source_hash": record["source_hash"],
        "human_review_status": None,
        "human_revenue_date": None,
        "human_reviewer_note": None,
    }


def line_item_rows(record: dict) -> list[dict]:
    rows = []
    for line_no, item in enumerate(record["line_items"] or [], start=1):
        rows.append({
            "record_id": record["record_id"],
            "line_no": line_no,
            **{key: item[key] for key in LINE_ITEM_COLUMNS[2:]},
        })
    return rows


def _write_csv(path: Path, columns: list[str], rows: Iterable[dict]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding=CSV_ENCODING, newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({key: _cell(row[key]) for key in columns})
    return path


def write_csv(path: Path, columns: list[str], rows: Iterable[dict]) -> Path:
    """BOM 付き UTF-8・CSV インジェクション対策・上書き禁止で CSV を書く（汎用）。"""
    return _write_csv(path, columns, rows)


def write_review_csv(records: list[dict], path: Path) -> Path:
    """1 行 = 1 sales_record の確認用 CSV。"""
    return _write_csv(path, RECORD_COLUMNS, (record_to_row(r) for r in records))


def write_line_items_csv(records: list[dict], path: Path) -> Path:
    """1 行 = 1 明細の確認用 CSV（record_id で対応付け）。"""
    rows = [row for record in records for row in line_item_rows(record)]
    return _write_csv(path, LINE_ITEM_COLUMNS, rows)

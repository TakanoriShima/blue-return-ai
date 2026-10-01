"""年間処理の入口（全件再生成方式）。

    python -m blue_return_ai.annual --year 2026 [--config data/config/annual.json] [--out-root data/output/annual]

- 実行のたびに data/output/annual/<run-id>/ を新しく作り、すべての成果物をそこに出力する。
  既存の出力先は上書きしない。
- 設定ファイルの processors に従い、売上 Excel（テンプレート）・カード CSV・銀行 CSV を取り込み、
  sales.csv / sales_line_items.csv / card_transactions.csv / bank_transactions.csv /
  unresolved_items.csv / documents_inventory.csv / manifest.json と、人間確認用の
  sales_summary_<年>_provisional.xlsx（売上・請求書集計（暫定））を出力する。
  照合・仕訳・集計は次の段階で追加する。
- 画面には件数だけを表示し、実ファイル名・実データの値は表示しない。
  資料一覧の CSV には元ファイルの相対パスを記録するが、出力先は data/ 配下に限る。

人間の判断（data/decisions/documents.csv）
    sha256,status,note
    <ファイルの SHA-256>,manual,<メモ>
  status には manual / cross_check_only / unsupported / unresolved / imported を指定できる。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sys
import uuid
from collections import Counter
from datetime import datetime
from pathlib import Path

from .cli import REPO_ROOT, UnsafeOutputDirError, check_out_dir
from .config import CATEGORIES, AnnualConfig, ConfigError, load_config
from .annual_processors import (
    BANK_COLUMNS,
    CARD_COLUMNS,
    SALES_COLUMNS,
    UNRESOLVED_COLUMNS,
    apply_accounting_amounts,
    process_documents,
    resolve_processors,
    sales_rows,
    summarize,
    transaction_rows,
    unresolved_rows,
)
from . import sales_workbook
from .output import write_csv, write_line_items_csv

MIN_YEAR = 2000
MAX_YEAR = 2099
MANIFEST_SCHEMA = "annual_manifest/2"
DEFAULT_CONFIG = REPO_ROOT / "data" / "config" / "annual.json"
DEFAULT_OUT_ROOT = REPO_ROOT / "data" / "output" / "annual"

DOCUMENT_STATUSES = ("imported", "manual", "cross_check_only", "unsupported", "unresolved")

# 取込対象とし得る拡張子。取込処理で取り込めたものは imported、取り込めない・未対応のものは unresolved のまま
PLANNED_EXTENSIONS = {
    "invoices": {".xlsx", ".pdf"},
    "bank": {".csv"},
    "credit_card": {".csv"},
    "household": {".csv", ".xlsx", ".pdf"},
}

INVENTORY_COLUMNS = [
    "document_key", "category", "extension", "size_bytes", "sha256",
    "status", "status_source", "duplicate_of", "relative_path", "note",
]

OUTPUT_FILES = [
    "sales.csv", "sales_line_items.csv", "card_transactions.csv", "bank_transactions.csv",
    "unresolved_items.csv", "documents_inventory.csv", "manifest.json",
]


def output_files(year: int) -> list[str]:
    """実行フォルダに作るファイル（人間確認用の売上 Excel を含む）。"""
    return [*OUTPUT_FILES, sales_workbook.workbook_file_name(year)]


class AnnualError(Exception):
    """年間処理の誤り。メッセージに実データの値・ファイル名を含めない。"""


def validate_year(year: object) -> int:
    if isinstance(year, bool) or not isinstance(year, int) or not MIN_YEAR <= year <= MAX_YEAR:
        raise AnnualError(f"年は {MIN_YEAR}〜{MAX_YEAR} の整数で指定してください")
    return year


def new_run_id(year: int, now: datetime | None = None) -> str:
    now = now or datetime.now()
    return f"{year}_{now:%Y%m%d-%H%M%S}_{uuid.uuid4().hex[:6]}"


def create_run_dir(out_root: Path, run_id: str) -> Path:
    out_root = check_out_dir(out_root)
    run_dir = out_root / run_id
    out_root.mkdir(parents=True, exist_ok=True)
    try:
        run_dir.mkdir()  # 既に存在すれば FileExistsError（上書きしない）
    except FileExistsError:
        raise AnnualError("同じ実行 ID の出力先が既に存在します（上書きしません）") from None
    return run_dir


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_document_decisions(decisions_dir: Path | None) -> dict[str, dict]:
    """data/decisions/documents.csv（sha256,status,note）を読む。無ければ空。"""
    if decisions_dir is None:
        return {}
    path = decisions_dir / "documents.csv"
    if not path.exists():
        return {}
    decisions: dict[str, dict] = {}
    with path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None or not {"sha256", "status"} <= set(reader.fieldnames):
            raise AnnualError("documents.csv には sha256 と status の列が必要です")
        for line, row in enumerate(reader, start=2):
            sha = (row.get("sha256") or "").strip().lower()
            status = (row.get("status") or "").strip()
            if len(sha) != 64 or any(c not in "0123456789abcdef" for c in sha):
                raise AnnualError(f"documents.csv の {line} 行目: sha256 が正しくありません")
            if status not in DOCUMENT_STATUSES:
                raise AnnualError(f"documents.csv の {line} 行目: status が正しくありません")
            decisions[sha] = {"status": status, "note": (row.get("note") or "").strip()}
    return decisions


def build_inventory(config: AnnualConfig, repo_root: Path,
                    decisions: dict[str, dict] | None = None) -> list[dict]:
    """設定された各フォルダのファイルを一覧化する（ファイルの中身は解析しない）。"""
    decisions = decisions or {}
    rows: list[dict] = []
    first_seen: dict[str, str] = {}
    for category in CATEGORIES:
        directory = config.sources.get(category)
        if directory is None or not directory.is_dir():
            continue
        for path in sorted(p for p in directory.rglob("*") if p.is_file()):
            resolved = path.resolve()
            if not resolved.is_relative_to(directory):
                continue  # リンク等でフォルダ外を指すものは対象外
            sha = _sha256_file(path)
            extension = path.suffix.lower()
            note = ""
            if path.name.startswith("~$"):
                status, note = "unsupported", "temporary_file"
            elif extension in PLANNED_EXTENSIONS[category]:
                status = "unresolved"
            else:
                status = "unsupported"
            status_source = "program"
            if sha in decisions:
                status = decisions[sha]["status"]
                note = decisions[sha]["note"] or note
                status_source = "human"
            document_key = f"document-{sha[:16]}"
            rows.append({
                "document_key": document_key,
                "category": category,
                "extension": extension or "(none)",
                "size_bytes": path.stat().st_size,
                "sha256": sha,
                "status": status,
                "status_source": status_source,
                "duplicate_of": first_seen.get(sha, ""),
                "relative_path": resolved.relative_to(repo_root.resolve()).as_posix(),
                "note": note,
            })
            first_seen.setdefault(sha, document_key)
    return rows


def run(year: int, config_path: Path, out_root: Path, repo_root: Path = REPO_ROOT) -> dict:
    year = validate_year(year)
    config = load_config(config_path, repo_root)
    processors = resolve_processors(config)
    decisions = load_document_decisions(config.decisions_dir)
    run_id = new_run_id(year)
    run_dir = create_run_dir(out_root, run_id)

    inventory = build_inventory(config, repo_root, decisions)
    processed = process_documents(inventory, processors, repo_root, year)
    accounting = config.accounting.get(year)
    apply_accounting_amounts(processed, accounting)
    unresolved = unresolved_rows(processed)
    counts = summarize(processed, unresolved)

    write_csv(run_dir / "sales.csv", SALES_COLUMNS, sales_rows(processed))
    write_line_items_csv([item["record"] for item in processed.sales], run_dir / "sales_line_items.csv")
    write_csv(run_dir / "card_transactions.csv", CARD_COLUMNS,
              transaction_rows(processed.card, CARD_COLUMNS))
    write_csv(run_dir / "bank_transactions.csv", BANK_COLUMNS,
              transaction_rows(processed.bank, BANK_COLUMNS))
    write_csv(run_dir / "unresolved_items.csv", UNRESOLVED_COLUMNS, unresolved)
    write_csv(run_dir / "documents_inventory.csv", INVENTORY_COLUMNS, inventory)
    workbook_name = sales_workbook.workbook_file_name(year)
    created_at = datetime.now().isoformat(timespec="seconds")
    sales_workbook.write_workbook(run_dir / workbook_name, sales_workbook.build_workbook(
        sales=sales_rows(processed),
        line_items=[
            {**line, "record_id": item["record"]["record_id"], "line_no": number}
            for item in processed.sales
            for number, line in enumerate(item["record"]["line_items"] or [], start=1)
        ],
        issues=[i for i in unresolved if i["source_type"] == "sales"],
        invoice_documents=Counter(r["status"] for r in inventory if r["category"] == "invoices"),
        year=year, run_id=run_id, created_at=created_at, accounting=accounting,
    ))

    by_status = Counter(row["status"] for row in inventory)
    by_category = {
        category: dict(Counter(r["status"] for r in inventory if r["category"] == category))
        for category in CATEGORIES
    }
    manifest = {
        "schema": MANIFEST_SCHEMA,
        "run_id": run_id,
        "year": year,
        "created_at": created_at,
        "config_sha256": config.sha256,
        "sources": {
            category: {
                "configured": category in config.sources,
                "exists": category in config.sources and config.sources[category].is_dir(),
            }
            for category in CATEGORIES
        },
        "processors": {
            category: config.processors[category]["id"] if category in config.processors
            else "not_configured"
            for category in CATEGORIES
        },
        "documents": {
            "total": len(inventory),
            "duplicates": sum(1 for r in inventory if r["duplicate_of"]),
            "by_status": {status: by_status.get(status, 0) for status in DOCUMENT_STATUSES},
            "by_category": by_category,
        },
        "accounting": accounting,  # 年度の会計設定（未設定なら null）
        "processing": counts,
        "outputs": output_files(year),
        "notes": [
            "暫定実行です。売上と入金の照合、カードと引落の照合、仕訳、集計はまだ実装されていません。",
            "売上計上日（revenue_date）は未確定です。取引の分類（事業・私用等）はすべて未判断です。",
            "出力 CSV には実データ（documents_inventory.csv は元ファイルのパス）が含まれます。外部 AI に渡さないでください。",
        ],
    }
    with (run_dir / "manifest.json").open("x", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return {"run_dir": run_dir, "manifest": manifest}


def format_summary(manifest: dict) -> str:
    documents = manifest["documents"]
    lines = [
        f"実行 ID: {manifest['run_id']}（{manifest['year']} 年・暫定）",
        f"資料: 合計 {documents['total']} 件、同一内容の重複 {documents['duplicates']} 件",
        "  状態別: " + ", ".join(f"{k}={v}" for k, v in documents["by_status"].items()),
    ]
    for category, counts in documents["by_category"].items():
        source = manifest["sources"][category]
        state = "フォルダなし" if not source["exists"] else (
            ", ".join(f"{k}={v}" for k, v in counts.items()) or "0 件")
        if not source["configured"]:
            state = "未設定"
        lines.append(f"  {category}: {state}")
    lines.append("取込処理: " + ", ".join(f"{k}={v}" for k, v in manifest["processors"].items()))
    processing = manifest["processing"]
    for category, counts in processing["files"].items():
        lines.append(f"  {category} の資料: " + ", ".join(f"{k}={v}" for k, v in counts.items()))
    sales = processing["sales"]
    lines.append(f"売上: {sales['records']} 件（要確認 {sales['needs_review']} 件）")
    lines.append("  帳票パターン別: " + (
        ", ".join(f"{k}={v}" for k, v in sales["by_template_pattern"].items()) or "なし"))
    for label, key in (("カード取引", "card"), ("銀行取引", "bank")):
        c = processing[key]
        lines.append(f"{label}: {c['transactions']} 件（要確認 {c['needs_review']} 件、"
                     f"重複として除外 {c['duplicate_transactions']} 件、"
                     f"取引として読まなかった行 {c['rows_not_transaction']} 件）")
    unresolved = processing["unresolved"]
    lines.append(f"要確認一覧: {unresolved['total']} 件（コード別）")
    for code, count in unresolved["by_code"].items():
        lines.append(f"  {code}: {count}")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m blue_return_ai.annual",
        description="年間の会計資料から暫定データを全件再生成する（売上・カード・銀行の取込まで）")
    parser.add_argument("--year", type=int, required=True, help="対象年（例: 2026）")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG,
                        help="設定ファイル（既定: data/config/annual.json）")
    parser.add_argument("--out-root", type=Path, default=DEFAULT_OUT_ROOT,
                        help="出力先の親フォルダ（既定: data/output/annual）")
    args = parser.parse_args(argv)
    try:
        result = run(args.year, args.config, args.out_root, repo_root=REPO_ROOT)
    except (AnnualError, ConfigError, UnsafeOutputDirError) as error:
        print(f"エラー: {error}", file=sys.stderr)
        return 2
    except Exception as error:  # noqa: BLE001 - 実データを含む可能性があるため詳細は表示しない
        print(f"エラー: 処理に失敗しました: {type(error).__name__}", file=sys.stderr)
        return 1
    print(format_summary(result["manifest"]))
    print(f"出力先: data/output/annual 配下の {result['run_dir'].name}")
    print(f"人間確認用: {sales_workbook.workbook_file_name(result['manifest']['year'])}（売上・請求書集計（暫定））")
    print("出力 CSV・Excel には実データが含まれます。外部 AI に渡さないでください（共有は上の件数のみ）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

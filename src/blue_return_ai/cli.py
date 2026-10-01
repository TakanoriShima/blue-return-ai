"""コマンドライン実行。

    python -m blue_return_ai.cli <Excelファイル...> [--template NAME] [--out-dir DIR] [--target-year YYYY]

- テンプレートは --template で明示する（既定: template_a）。自動判定はしない。
- 出力先の既定値は data/output（Git 管理対象外）。
  リポジトリ内では data/ 配下以外への出力を拒否する。
- 画面には実データの値・元ファイル名を表示しない
  （件数・record_id・review_status・警告コードのみ）。
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from . import template_a, template_b
from .output import (
    find_records_by_source_hash,
    write_line_items_csv,
    write_record_json,
    write_review_csv,
)
from .sales_record import build_sales_record
from .validation import validate

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_OUT_DIR = REPO_ROOT / "data" / "output"

TEMPLATES = {
    template_a.SOURCE_NAME: template_a,
    template_b.SOURCE_NAME: template_b,
}


class UnsafeOutputDirError(Exception):
    pass


def check_out_dir(out_dir: Path, repo_root: Path = REPO_ROOT) -> Path:
    """リポジトリ内の data/ 以外（Git 管理対象になり得る場所）への出力を拒否する。"""
    resolved = out_dir.resolve()
    if resolved.is_relative_to(repo_root) and not resolved.is_relative_to(repo_root / "data"):
        raise UnsafeOutputDirError(
            "リポジトリ内では data/ 配下以外へ出力できません（実データが Git 管理対象になるのを防ぐため）"
        )
    return resolved


def process_file(path: Path, template_name: str, target_year: int | None) -> dict:
    template = TEMPLATES[template_name]
    source_hash, extracted, issues = template.read(path)
    record = build_sales_record(
        extracted=extracted,
        issues=issues,
        source_hash=source_hash,
        source_name=template.SOURCE_NAME,
    )
    return validate(record, target_year=target_year)


def run(inputs: list[Path], template_name: str, out_dir: Path,
        target_year: int | None = None) -> int:
    out_dir = check_out_dir(out_dir)
    records_dir = out_dir / "records"
    review_dir = out_dir / "review"

    records: list[dict] = []
    seen_hashes: set[str] = set()
    failed = 0
    total = len(inputs)

    for index, path in enumerate(inputs, start=1):
        label = f"[{index}/{total}]"
        try:
            record = process_file(path, template_name, target_year)
        except Exception as error:  # noqa: BLE001 - 実データを含む可能性があるため詳細は表示しない
            failed += 1
            print(f"{label} 読み取りに失敗しました: {type(error).__name__}")
            continue

        existing = find_records_by_source_hash(records_dir, record["source_hash"])
        if existing or record["source_hash"] in seen_hashes:
            print(f"{label} 既存取込候補（source_hash が一致）のため登録しません: "
                  f"既存 record_id={','.join(existing) or '同一実行内で重複'}")
            continue
        seen_hashes.add(record["source_hash"])

        write_record_json(record, records_dir)
        records.append(record)
        codes = ",".join(sorted({w["code"] for w in record["warnings"]})) or "-"
        print(f"{label} record_id={record['record_id']} "
              f"review_status={record['review_status']} warnings={codes}")

    if records:
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        write_review_csv(records, review_dir / f"{stamp}_records.csv")
        write_line_items_csv(records, review_dir / f"{stamp}_line_items.csv")

    needs_review = sum(r["review_status"] == "needs_review" for r in records)
    print(f"完了: 登録 {len(records)} 件（うち要確認 {needs_review} 件）、"
          f"スキップ {total - len(records) - failed} 件、失敗 {failed} 件")
    return 1 if failed else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Excel 請求書を共通 sales_record へ変換・検証する")
    parser.add_argument("inputs", nargs="+", type=Path, help="Excel 請求書（.xlsx）")
    parser.add_argument("--template", default=template_a.SOURCE_NAME, choices=sorted(TEMPLATES),
                        help="適用するテンプレート（既定: template_a。自動判定はしない）")
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR,
                        help="出力先（既定: data/output）")
    parser.add_argument("--target-year", type=int, default=None,
                        help="対象年。指定すると請求日が範囲外の場合に警告する")
    args = parser.parse_args(argv)

    try:
        return run(args.inputs, args.template, args.out_dir, args.target_year)
    except UnsafeOutputDirError as error:
        print(f"エラー: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())

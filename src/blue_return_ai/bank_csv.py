"""銀行入出金明細 CSV の取込（書式 bank_csv_a。docs/data_model.md §17.3）。

- 出金額（withdrawal_amount）と入金額（deposit_amount）は別の項目として保持する。
  空欄は null とし、0 に変換しない（0 と書かれていれば 0）。
- 摘要から売上入金・経費・カード引落・私用・振替を推測しない。分類ルール・人間の判断が無いものは unknown（要確認）。
- 列数が違う行・日付を解釈できない行は取引として取り込まず、要確認一覧に行番号だけを出す。
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path

from .csv_table import (
    CsvFormat,
    fit_columns,
    is_blank,
    parse_csv_amount,
    parse_csv_date,
    read_table,
    text_or_none,
)
from .issues import make_issue
from .keys import transaction_source_key

# 人間が構造調査で確認した列構成（公開コードには金融機関名を書かない）
BANK_CSV_A = CsvFormat(
    format_id="bank_csv_a",
    encoding="cp932",
    header_row=13,
    columns=(
        "明細通番",
        "日付",
        "お引出金額",
        "お預入金額",
        "残高",
        "お取引内容",
    ),
)
FORMATS = {BANK_CSV_A.format_id: BANK_CSV_A}


def parse_bank_csv(path: Path, fmt: CsvFormat, *, document_key: str,
                   target_year: int | None = None) -> tuple[list[dict], list[dict]]:
    """(取引の一覧, 取り込まなかった行の要確認) を返す。CsvFormatError は呼び出し側で扱う。"""
    source_hash, rows = read_table(path, fmt)
    transactions: list[dict] = []
    row_issues: list[dict] = []
    seen: Counter = Counter()

    for row in rows:
        if is_blank(row.fields):
            continue
        fields = fit_columns(row.fields, len(fmt.columns))
        if fields is None:
            row_issues.append({"row_number": row.row_number, "code": "ROW_NOT_TRANSACTION",
                               "field": "row"})
            continue
        sequence, date_text, withdrawal_text, deposit_text, balance_text, description = fields

        transaction_date, date_code = parse_csv_date(date_text)
        if transaction_date is None:
            row_issues.append({"row_number": row.row_number, "code": "ROW_NOT_TRANSACTION",
                               "field": "transaction_date"})
            continue

        issues: list[dict] = []
        withdrawal, withdrawal_code = parse_csv_amount(withdrawal_text)
        deposit, deposit_code = parse_csv_amount(deposit_text)
        balance, balance_code = parse_csv_amount(balance_text)
        for code, field in ((date_code, "transaction_date"), (withdrawal_code, "withdrawal_amount"),
                            (deposit_code, "deposit_amount"), (balance_code, "balance")):
            if code:
                issues.append(make_issue(code, field))
        if (withdrawal is None and deposit is None
                and withdrawal_code is None and deposit_code is None):
            issues.append(make_issue("BANK_AMOUNT_MISSING", "withdrawal_amount"))
        if withdrawal and deposit:  # どちらも 0 以外の値がある
            issues.append(make_issue("BANK_BOTH_AMOUNTS", "deposit_amount"))
        for value, field in ((withdrawal, "withdrawal_amount"), (deposit, "deposit_amount")):
            if value is not None and value < 0:
                issues.append(make_issue("AMOUNT_NEGATIVE", field))
        if balance is None and balance_code is None:
            issues.append(make_issue("FIELD_MISSING", "balance"))
        if target_year is not None and int(transaction_date[:4]) != target_year:
            issues.append(make_issue("DATE_OUT_OF_TARGET_YEAR", "transaction_date"))
        issues.append(make_issue("CLASSIFICATION_UNKNOWN", "classification"))

        # 明細通番は、ダウンロードの単位ごとに振り直される可能性があるためキーに含めない（§17.4）
        content = [transaction_date, withdrawal_text.strip(), deposit_text.strip(),
                   balance_text.strip(), text_or_none(description)]
        seen[tuple(content)] += 1
        transactions.append({
            "source_key": transaction_source_key("bank", content, seen[tuple(content)]),
            "document_key": document_key,
            "source_hash": source_hash,
            "format_id": fmt.format_id,
            "row_number": row.row_number,
            "record_status": "active",
            "duplicate_of_document": None,
            "sequence": text_or_none(sequence),
            "transaction_date": transaction_date,
            "withdrawal_amount": withdrawal,
            "deposit_amount": deposit,
            "balance": balance,
            "description": text_or_none(description),
            "classification": "unknown",
            "account_category": None,
            "review_status": "needs_review",
            "warnings": issues,
        })
    return transactions, row_issues

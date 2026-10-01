"""カード利用明細 CSV の取込（書式 card_csv_a。docs/data_model.md §17.2）。

- 経費計上日の候補は「利用日（use_date）」。支払金額や支払月を経費の発生日として扱わない。
- 利用金額（amount）と今回の支払金額（payment_amount）は別の項目として保持する。
- 利用先名から事業・私用を推測しない。分類ルール・人間の判断が無いものは unknown（要確認）。
- 返金等の負の金額も読み込む（要確認の警告を付ける）。
- 列数が違う行・利用日を解釈できない行（合計行等）は取引として取り込まず、要確認一覧に行番号だけを出す。
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
CARD_CSV_A = CsvFormat(
    format_id="card_csv_a",
    encoding="cp932",
    header_row=5,
    columns=(
        "ご利用区分",
        "ご利用日",
        "ご利用者区分",
        "ご利用店",
        "ポイント対象",
        "今回回数",
        "ご利用金額",
        "今回のお支払い金額",
        "備考",
    ),
)
FORMATS = {CARD_CSV_A.format_id: CARD_CSV_A}

_SINGLE_PAYMENT = {"1", "1回", "01"}


def parse_card_csv(path: Path, fmt: CsvFormat, *, document_key: str,
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
        (usage_category, use_date_text, user_category, merchant, point_target,
         installment, amount_text, payment_text, note) = fields

        use_date, date_code = parse_csv_date(use_date_text)
        if use_date is None:
            # 合計行など、利用日の無い行は取引として扱わない
            row_issues.append({"row_number": row.row_number, "code": "ROW_NOT_TRANSACTION",
                               "field": "use_date"})
            continue

        issues: list[dict] = []
        amount, amount_code = parse_csv_amount(amount_text)
        payment_amount, payment_code = parse_csv_amount(payment_text)
        for code, field in ((date_code, "use_date"), (amount_code, "amount"),
                            (payment_code, "payment_amount")):
            if code:
                issues.append(make_issue(code, field))
        if amount is None and amount_code is None:
            issues.append(make_issue("FIELD_MISSING", "amount"))
        if amount is not None and amount < 0:
            issues.append(make_issue("CARD_NEGATIVE_AMOUNT", "amount"))
        installment_text = text_or_none(installment)
        if (installment_text is not None and installment_text not in _SINGLE_PAYMENT) or (
                amount is not None and payment_amount is not None and amount != payment_amount):
            issues.append(make_issue("CARD_INSTALLMENT", "payment_amount"))
        if target_year is not None and int(use_date[:4]) != target_year:
            issues.append(make_issue("DATE_OUT_OF_TARGET_YEAR", "use_date"))
        issues.append(make_issue("CLASSIFICATION_UNKNOWN", "business_classification"))

        content = [use_date, text_or_none(merchant), amount_text.strip(), payment_text.strip(),
                   installment_text, text_or_none(usage_category), text_or_none(user_category)]
        seen[tuple(content)] += 1
        transactions.append({
            "source_key": transaction_source_key("card", content, seen[tuple(content)]),
            "document_key": document_key,
            "source_hash": source_hash,
            "format_id": fmt.format_id,
            "row_number": row.row_number,
            "record_status": "active",
            "duplicate_of_document": None,
            "use_date": use_date,
            "usage_category": text_or_none(usage_category),
            "user_category": text_or_none(user_category),
            "merchant": text_or_none(merchant),
            "point_target": text_or_none(point_target),
            "installment_count": installment_text,
            "amount": amount,
            "payment_amount": payment_amount,
            "note": text_or_none(note),
            "business_classification": "unknown",
            "account_category": None,
            "review_status": "needs_review",
            "warnings": issues,
        })
    return transactions, row_issues

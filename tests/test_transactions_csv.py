"""カード CSV（card_csv_a）・銀行 CSV（bank_csv_a）取込のテスト。すべて完全な架空データ。"""

from __future__ import annotations

import pytest

from blue_return_ai.bank_csv import BANK_CSV_A, parse_bank_csv
from blue_return_ai.card_csv import CARD_CSV_A, parse_card_csv
from blue_return_ai.csv_table import CsvFormatError

CARD_HEADER = ",".join(CARD_CSV_A.columns)
BANK_HEADER = ",".join(BANK_CSV_A.columns)

# ヘッダーより前の説明行（架空）。取込処理は内容を解析しない
CARD_PREAMBLE = ["架空カード ご利用明細", "会員番号,****-****-****-0000", "お支払日,2026/07/27", ""]
BANK_PREAMBLE = ["架空銀行 入出金明細", "店番,000", "口座番号,0000000", "口座名義,カクウ タロウ",
                 "照会期間,2026/06/01-2026/06/30", "", "", "", "", "", "", ""]

CARD_ROWS = [
    "1回払い,2026/06/03,本人,架空書店,○,1,1200,1200,",
    "1回払い,2026/06/03,本人,架空書店,○,1,1200,1200,",  # 同じ内容の正当な別取引
    "",
    "1回払い,2026/06/10,家族,架空電器店,○,,-3000,-3000,返品",
    "分割,2026/05/20,本人,架空家具店,,2,\"30,000\",\"10,000\",",
    "1回払い,2026/06/15,本人,架空カフェ,,1,,,",
    "1回払い,2025/12/28,本人,架空交通,,1,800,800,",
    "合計,,,,,,\"39,000\",\"10,000\",",
    "列数の違う行,1",
]
BANK_ROWS = [
    "1,2026/06/01,,\"50,000\",\"1,050,000\",架空入金 カ）カクウ",
    "2,2026/06/02,3000,,\"1,047,000\",架空公共料金",
    "3,2026/06/02,3000,,\"1,044,000\",架空公共料金",  # 同日同額の正当な別取引（残高が異なる）
    "4,2026/06/03,0,,\"1,044,000\",架空手数料",
    "5,2026/06/04,,,\"1,044,000\",架空不明",
    "6,20260605,1000,,,架空日付8桁",
    "",
    "合計,,,,,",
]


def write_csv(path, preamble, header, rows, encoding="cp932"):
    path.write_bytes(("\r\n".join([*preamble, header, *rows]) + "\r\n").encode(encoding))
    return path


@pytest.fixture
def card_file(tmp_path):
    return write_csv(tmp_path / "card.csv", CARD_PREAMBLE, CARD_HEADER, CARD_ROWS)


@pytest.fixture
def bank_file(tmp_path):
    return write_csv(tmp_path / "bank.csv", BANK_PREAMBLE, BANK_HEADER, BANK_ROWS)


def codes(transaction):
    return {(w["code"], w["field"]) for w in transaction["warnings"]}


# ---------------------------------------------------------------- card

def test_card_header_is_row_5_and_encoding_cp932():
    assert CARD_CSV_A.header_row == 5
    assert CARD_CSV_A.encoding == "cp932"


def test_card_transactions(card_file):
    transactions, row_issues = parse_card_csv(card_file, CARD_CSV_A, document_key="document-x",
                                              target_year=2026)
    assert len(transactions) == 6
    first = transactions[0]
    assert first["use_date"] == "2026-06-03"
    assert first["merchant"] == "架空書店"
    assert first["usage_category"] == "1回払い"
    assert first["user_category"] == "本人"
    assert first["amount"] == 1200
    assert first["payment_amount"] == 1200
    assert first["installment_count"] == "1"
    assert first["note"] is None
    assert first["business_classification"] == "unknown"
    assert first["account_category"] is None
    assert first["review_status"] == "needs_review"
    assert ("CLASSIFICATION_UNKNOWN", "business_classification") in codes(first)
    assert first["row_number"] == 6
    # 合計行・列数の違う行は取引にしない
    assert [(i["row_number"], i["code"]) for i in row_issues] == [
        (13, "ROW_NOT_TRANSACTION"), (14, "ROW_NOT_TRANSACTION")]


def test_card_same_content_in_one_file_is_kept(card_file):
    transactions, _ = parse_card_csv(card_file, CARD_CSV_A, document_key="d")
    assert transactions[0]["source_key"] != transactions[1]["source_key"]


def test_card_negative_installment_blank_and_year(card_file):
    transactions, _ = parse_card_csv(card_file, CARD_CSV_A, document_key="d", target_year=2026)
    refund, installment, blank, last_year = transactions[2:6]
    assert refund["amount"] == -3000
    assert ("CARD_NEGATIVE_AMOUNT", "amount") in codes(refund)
    assert installment["amount"] == 30000
    assert installment["payment_amount"] == 10000
    assert ("CARD_INSTALLMENT", "payment_amount") in codes(installment)
    assert blank["amount"] is None
    assert blank["payment_amount"] is None
    assert ("FIELD_MISSING", "amount") in codes(blank)
    assert last_year["use_date"] == "2025-12-28"
    assert ("DATE_OUT_OF_TARGET_YEAR", "use_date") in codes(last_year)


def test_card_use_date_is_expense_date_not_payment_month(card_file):
    transactions, _ = parse_card_csv(card_file, CARD_CSV_A, document_key="d")
    assert {t["use_date"][:7] for t in transactions} == {"2026-06", "2026-05", "2025-12"}
    assert all("payment_date" not in t for t in transactions)


def test_card_key_does_not_contain_values(card_file):
    transactions, _ = parse_card_csv(card_file, CARD_CSV_A, document_key="d")
    for transaction in transactions:
        key = transaction["source_key"]
        assert key.startswith("card-1-")
        for value in ("架空", "1200", "2026"):
            assert value not in key


def test_card_header_mismatch_rejects_whole_file(tmp_path):
    path = write_csv(tmp_path / "card.csv", CARD_PREAMBLE, "日付,店名,金額", CARD_ROWS)
    with pytest.raises(CsvFormatError) as error:
        parse_card_csv(path, CARD_CSV_A, document_key="d")
    assert "架空" not in str(error.value)


def test_card_wrong_encoding_is_rejected(tmp_path):
    path = write_csv(tmp_path / "card.csv", CARD_PREAMBLE, CARD_HEADER, CARD_ROWS, encoding="utf-16")
    with pytest.raises(CsvFormatError):
        parse_card_csv(path, CARD_CSV_A, document_key="d")


def test_card_header_option_override(tmp_path):
    path = write_csv(tmp_path / "card.csv", ["架空"], CARD_HEADER, CARD_ROWS[:1])
    fmt = CARD_CSV_A.with_overrides({"header_row": 2})
    transactions, _ = parse_card_csv(path, fmt, document_key="d")
    assert len(transactions) == 1
    with pytest.raises(ValueError):
        CARD_CSV_A.with_overrides({"header_row": 0})
    with pytest.raises(ValueError):
        CARD_CSV_A.with_overrides({"delimiter": "tab"})


# ---------------------------------------------------------------- bank

def test_bank_header_is_row_13_and_encoding_cp932():
    assert BANK_CSV_A.header_row == 13
    assert BANK_CSV_A.encoding == "cp932"


def test_bank_transactions(bank_file):
    transactions, row_issues = parse_bank_csv(bank_file, BANK_CSV_A, document_key="d",
                                              target_year=2026)
    assert len(transactions) == 6
    deposit, withdrawal, withdrawal_2, zero, missing, compact = transactions
    assert deposit["sequence"] == "1"
    assert deposit["transaction_date"] == "2026-06-01"
    assert deposit["deposit_amount"] == 50000
    assert deposit["withdrawal_amount"] is None
    assert deposit["balance"] == 1050000
    assert deposit["description"] == "架空入金 カ）カクウ"
    assert deposit["classification"] == "unknown"
    assert deposit["account_category"] is None
    assert ("CLASSIFICATION_UNKNOWN", "classification") in codes(deposit)
    assert withdrawal["withdrawal_amount"] == 3000
    assert withdrawal["deposit_amount"] is None
    # 0 は 0、空欄は null
    assert zero["withdrawal_amount"] == 0
    assert zero["deposit_amount"] is None
    assert ("BANK_AMOUNT_MISSING", "withdrawal_amount") in codes(missing)
    assert compact["transaction_date"] == "2026-06-05"
    assert compact["balance"] is None
    assert ("FIELD_MISSING", "balance") in codes(compact)
    assert [(i["row_number"], i["code"]) for i in row_issues] == [(21, "ROW_NOT_TRANSACTION")]


def test_bank_same_day_same_amount_are_kept(bank_file):
    transactions, _ = parse_bank_csv(bank_file, BANK_CSV_A, document_key="d")
    assert transactions[1]["source_key"] != transactions[2]["source_key"]


def test_bank_key_ignores_sequence(tmp_path):
    first = write_csv(tmp_path / "a.csv", BANK_PREAMBLE, BANK_HEADER, BANK_ROWS[:2])
    renumbered = ["101" + row[1:] for row in BANK_ROWS[:2]]
    second = write_csv(tmp_path / "b.csv", BANK_PREAMBLE, BANK_HEADER, renumbered)
    keys_a = [t["source_key"] for t in parse_bank_csv(first, BANK_CSV_A, document_key="a")[0]]
    keys_b = [t["source_key"] for t in parse_bank_csv(second, BANK_CSV_A, document_key="b")[0]]
    assert keys_a == keys_b


def test_bank_both_amounts_and_negative(tmp_path):
    rows = ["1,2026/06/01,100,200,1000,架空", "2,2026/06/02,-100,,1100,架空"]
    path = write_csv(tmp_path / "bank.csv", BANK_PREAMBLE, BANK_HEADER, rows)
    both, negative = parse_bank_csv(path, BANK_CSV_A, document_key="d")[0]
    assert ("BANK_BOTH_AMOUNTS", "deposit_amount") in codes(both)
    assert ("AMOUNT_NEGATIVE", "withdrawal_amount") in codes(negative)


def test_bank_preamble_is_not_parsed(tmp_path):
    # 説明行の列数・内容が変わっても、ヘッダー以降だけを読む
    preamble = ["架空,説明,行,が,たくさん,あり,ます,。"] * 12
    path = write_csv(tmp_path / "bank.csv", preamble, BANK_HEADER, BANK_ROWS[:1])
    transactions, row_issues = parse_bank_csv(path, BANK_CSV_A, document_key="d")
    assert len(transactions) == 1
    assert row_issues == []


def test_bank_header_mismatch_rejects_whole_file(tmp_path):
    path = write_csv(tmp_path / "bank.csv", BANK_PREAMBLE[:-1], BANK_HEADER, BANK_ROWS)
    with pytest.raises(CsvFormatError):
        parse_bank_csv(path, BANK_CSV_A, document_key="d")

"""検証・抽出で検出した警告（docs/data_model.md §8.2）。

警告には実データの値を含めない。コード・項目名・重大度のみで表現する。
"""

from __future__ import annotations

ERROR = "error"
WARNING = "warning"

SEVERITY: dict[str, str] = {
    # 抽出（§11）
    "FORMULA_NO_CACHED_VALUE": WARNING,
    "TEMPLATE_LABEL_MISMATCH": ERROR,
    "QUANTITY_INVALID": WARNING,
    # 欠落（§8.2）
    "FIELD_MISSING": WARNING,
    # 金額（§5.6）
    "AMOUNT_NOT_INTEGER": ERROR,
    "AMOUNT_NEGATIVE": ERROR,
    "NET_TAX_GROSS_MISMATCH": ERROR,
    "LINE_ITEMS_SUM_MISMATCH": ERROR,
    "LINE_ITEM_AMOUNT_MISMATCH": WARNING,
    "TAX_RATE_MISMATCH": WARNING,
    "WITHHOLDING_INCONSISTENT": ERROR,
    "EXPECTED_PAYMENT_MISMATCH": ERROR,
    "AMOUNT_UNUSUALLY_LARGE": WARNING,
    # 日付（§7）
    "DATE_INVALID": ERROR,
    "DATE_AMBIGUOUS": WARNING,
    "DATE_OUT_OF_TARGET_YEAR": WARNING,
    "DUE_BEFORE_DOCUMENT": ERROR,
    "DUE_TOO_FAR": WARNING,
    "PERIOD_REVERSED": ERROR,
    # 年間処理：資料・CSV（docs/data_model.md §17）
    "FILE_READ_FAILED": ERROR,
    "CSV_HEADER_MISMATCH": ERROR,
    "ROW_NOT_TRANSACTION": WARNING,
    # 年間処理：取引
    "CLASSIFICATION_UNKNOWN": WARNING,
    "DUPLICATE_CANDIDATE": WARNING,
    "CARD_NEGATIVE_AMOUNT": WARNING,
    "CARD_INSTALLMENT": WARNING,
    "BANK_AMOUNT_MISSING": ERROR,
    "BANK_BOTH_AMOUNTS": ERROR,
    # 年間処理：売上
    "REVENUE_DATE_UNCONFIRMED": WARNING,
}

# 要確認一覧（unresolved_items.csv）に出す説明。実データの値は含めない
MESSAGES: dict[str, str] = {
    "FORMULA_NO_CACHED_VALUE": "数式セルに計算結果（キャッシュ値）がありません",
    "TEMPLATE_LABEL_MISMATCH": "テンプレート定義のラベルが見つからない・重複している・位置が想定と異なります",
    "QUANTITY_INVALID": "数量を数値として解釈できません",
    "FIELD_MISSING": "値を取得できません（推測で補完していません）",
    "AMOUNT_NOT_INTEGER": "円単位の整数として解釈できない金額です",
    "AMOUNT_NEGATIVE": "金額が負です",
    "NET_TAX_GROSS_MISMATCH": "税抜金額・消費税額・税込総額が整合しません",
    "LINE_ITEMS_SUM_MISMATCH": "明細の合計が総額と一致しません",
    "LINE_ITEM_AMOUNT_MISMATCH": "明細の数量×単価が金額と一致しません",
    "TAX_RATE_MISMATCH": "消費税額が税率と整合しません",
    "WITHHOLDING_INCONSISTENT": "源泉徴収の状態と税額の組み合わせが不正です",
    "EXPECTED_PAYMENT_MISMATCH": "振込予定額が計算式と一致しません",
    "AMOUNT_UNUSUALLY_LARGE": "金額が設定した上限を超えています",
    "DATE_INVALID": "日付を解釈できません",
    "DATE_AMBIGUOUS": "日付が一意に解釈できません（年が無い等）",
    "DATE_OUT_OF_TARGET_YEAR": "日付が対象年の範囲外です（年をまたぐ取引の可能性）",
    "DUE_BEFORE_DOCUMENT": "支払期限が資料日付より前です",
    "DUE_TOO_FAR": "支払期限が資料日付から離れすぎています",
    "PERIOD_REVERSED": "業務期間の開始と終了が逆です",
    "FILE_READ_FAILED": "資料を読み取れませんでした",
    "CSV_HEADER_MISMATCH": "CSV のヘッダーが設定した書式と一致しません（資料全体を取り込んでいません）",
    "ROW_NOT_TRANSACTION": "取引として解釈できない行のため取り込んでいません（合計行・列数違い等）",
    "CLASSIFICATION_UNKNOWN": "事業・私用等の分類が未判断です",
    "DUPLICATE_CANDIDATE": "同じ日付・内容・金額の取引が別にあります（重複の可能性）",
    "CARD_NEGATIVE_AMOUNT": "利用金額が負です（返金・取消の可能性）",
    "CARD_INSTALLMENT": "分割・リボ等の可能性があります（複数回の支払明細に同じ利用が載る場合があります）",
    "BANK_AMOUNT_MISSING": "出金額・入金額のどちらも取得できません",
    "BANK_BOTH_AMOUNTS": "出金額と入金額の両方に値があります",
    "REVENUE_DATE_UNCONFIRMED": "売上計上日が未確定です（請求日を自動で計上日にしていません）",
}


def make_issue(code: str, field: str | None) -> dict:
    """警告を 1 件作る。未定義のコードは受け付けない。"""
    return {"code": code, "field": field, "severity": SEVERITY[code]}

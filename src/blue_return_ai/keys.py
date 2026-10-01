"""実行をまたいで安定するキー（source_key）。

- record_id（実行ごとの UUID）とは別のキー。人間の判断ファイル（data/decisions/）との紐づけに使う。
- 個人情報そのものをキーにしない。構成要素を SHA-256 でハッシュ化した値を使う。
- 売上資料のキーは「元ファイルの SHA-256 ＋ 資料内の位置」から作るため、元ファイルを持たない第三者には
  元の内容を推測できない。
"""

from __future__ import annotations

import hashlib

KEY_VERSION = "1"
_SEPARATOR = "\x1f"
_NAMESPACES = ("sales", "bank", "card", "household", "document")


def make_source_key(namespace: str, *parts: str) -> str:
    """namespace と構成要素から、決定的なキーを作る（例: sales-1-3f9a…）。"""
    if namespace not in _NAMESPACES:
        raise ValueError(f"namespace は {', '.join(_NAMESPACES)} のいずれかです")
    if not parts or any(not isinstance(p, str) or p == "" for p in parts):
        raise ValueError("キーの構成要素は空でない文字列で指定してください")
    canonical = _SEPARATOR.join((namespace, KEY_VERSION, *parts))
    digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    return f"{namespace}-{KEY_VERSION}-{digest[:32]}"


def transaction_source_key(namespace: str, fields: list[str | None], occurrence: int) -> str:
    """銀行・カード取引の source_key（docs/data_model.md §17.4）。

    取引の内容（日付・金額・摘要等の文字列）と、同じ内容の取引がそのファイル内で何回目に出たか（occurrence）から作る。
    ファイルの SHA-256 は含めないため、期間の重なる別ファイルに同じ取引があっても同じキーになる（重複除去に使う）。

    限界: 公開コードに秘密値を持たないため、日付・金額・摘要の候補が分かれば総当たりで一致を確かめられる。
    キーは data/ 配下の出力にだけ保存し、外部に共有しない。
    """
    if namespace not in ("bank", "card"):
        raise ValueError("namespace は bank または card です")
    if not isinstance(occurrence, int) or isinstance(occurrence, bool) or occurrence < 1:
        raise ValueError("occurrence は 1 以上の整数で指定してください")
    parts = [f"{index}={'' if value is None else value}" for index, value in enumerate(fields)]
    return make_source_key(namespace, *parts, f"occurrence={occurrence}")


def sales_source_key(source_hash: str, position: int = 1) -> str:
    """売上資料の source_key。position は資料内の何件目か（1 ファイル = 1 請求書なら 1）。"""
    if len(source_hash) != 64 or any(c not in "0123456789abcdef" for c in source_hash):
        raise ValueError("source_hash は SHA-256 の 16 進文字列（64 文字）で指定してください")
    if not isinstance(position, int) or isinstance(position, bool) or position < 1:
        raise ValueError("position は 1 以上の整数で指定してください")
    return make_source_key("sales", source_hash, str(position))

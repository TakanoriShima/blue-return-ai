"""年間処理の取込（売上 Excel・カード CSV・銀行 CSV）と、要確認一覧の作成（docs/data_model.md §17）。

- 取り込むのは、資料一覧で program が unresolved とした資料だけ。人間が判断ファイルで状態を指定した資料は取り込まない。
- 売上の自動取込は Excel のテンプレートだけ。PDF からは売上を作らない（Excel との二重計上を防ぐ）。
- 取引は source_key で重複を除く（期間の重なる CSV 等）。正当な同一内容の取引はファイル内の出現回数で区別する。
- どの処理も、実データの値・ファイル名を例外メッセージや画面に出さない。
"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from . import template_a, template_b, template_c, template_instructor
from .bank_csv import FORMATS as BANK_FORMATS
from .bank_csv import parse_bank_csv
from .card_csv import FORMATS as CARD_FORMATS
from .card_csv import parse_card_csv
from .config import AnnualConfig, ConfigError
from .csv_table import CsvFormatError
from .excel_reader import SheetNotFoundError, sheet_names
from .issues import MESSAGES, SEVERITY, make_issue
from .keys import sales_source_key
from .output import RECORD_COLUMNS, record_to_row
from .sales_record import build_sales_record
from .validation import validate

INVOICE_TEMPLATES = {
    template_a.SOURCE_NAME: template_a,
    template_b.SOURCE_NAME: template_b,
    template_c.SOURCE_NAME: template_c,
    template_instructor.SOURCE_NAME: template_instructor,
}
CSV_FORMATS = {"credit_card": CARD_FORMATS, "bank": BANK_FORMATS}
PROCESSOR_EXTENSIONS = {"invoices": ".xlsx", "credit_card": ".csv", "bank": ".csv"}

# テンプレートが共通モデルに割り当てずに返す確認用の値（template_c のみ。他のテンプレートでは空欄）
SALES_REFERENCE_FIELDS = ["template_pattern", "document_total", "document_subtotal", "withholding_base"]
# 会計用の売上金額（docs/data_model.md §17.7）。原資料の金額とは別の項目
ACCOUNTING_FIELDS = ["accounting_sales_amount", "accounting_sales_amount_basis"]
SALES_COLUMNS = ["source_key", "document_key", *RECORD_COLUMNS, "revenue_date", "revenue_date_status",
                 *SALES_REFERENCE_FIELDS, *ACCOUNTING_FIELDS]
CARD_COLUMNS = [
    "source_key", "document_key", "row_number", "record_status", "duplicate_of_document",
    "format_id", "use_date", "usage_category", "user_category", "merchant", "point_target",
    "installment_count", "amount", "payment_amount", "note",
    "business_classification", "account_category", "review_status", "warnings", "source_hash",
]
BANK_COLUMNS = [
    "source_key", "document_key", "row_number", "record_status", "duplicate_of_document",
    "format_id", "sequence", "transaction_date", "withdrawal_amount", "deposit_amount", "balance",
    "description", "classification", "account_category", "review_status", "warnings", "source_hash",
]
UNRESOLVED_COLUMNS = [
    "item_id", "source_type", "source_key", "document_key", "issue_code", "severity", "field",
    "message", "decision_status", "decision_value", "decision_note",
]


@dataclass
class ProcessResult:
    sales: list[dict] = field(default_factory=list)  # {"source_key", "document_key", "record"}
    card: list[dict] = field(default_factory=list)
    bank: list[dict] = field(default_factory=list)
    file_issues: list[dict] = field(default_factory=list)  # 資料単位・行単位の要確認
    stats: dict[str, Counter] = field(default_factory=lambda: defaultdict(Counter))


def resolve_processors(config: AnnualConfig) -> dict[str, object]:
    """設定の識別子を実際のテンプレート・書式に対応付ける。未知の識別子は ConfigError。"""
    resolved: dict[str, object] = {}
    for category, setting in config.processors.items():
        if category == "invoices":
            if any(name not in INVOICE_TEMPLATES for name in setting["ids"]):
                raise ConfigError("processors.invoices のテンプレート名が不明です")
            templates = [INVOICE_TEMPLATES[name] for name in setting["ids"]]
            if len({t.SHEET_NAME for t in templates}) != len(templates):
                raise ConfigError("processors.invoices に同じシート名のテンプレートを複数指定できません")
            resolved[category] = templates
            continue
        formats = CSV_FORMATS[category]
        if setting["id"] not in formats:
            raise ConfigError(f"processors.{category}.format が不明です")
        try:
            resolved[category] = formats[setting["id"]].with_overrides(setting["options"])
        except ValueError as error:
            raise ConfigError(f"processors.{category}.options: {error}") from None
    return resolved


def process_documents(inventory: list[dict], processors: dict[str, object],
                      repo_root: Path, year: int) -> ProcessResult:
    """資料一覧（inventory）の各行を取り込み、取込結果に応じて行の status / note を更新する。"""
    result = ProcessResult()
    for row in inventory:
        category = row["category"]
        if category not in processors:
            continue
        stats = result.stats[category]
        if row["status_source"] == "human":
            stats["skipped_by_decision"] += 1
            continue
        if row["status"] != "unresolved" or row["extension"] != PROCESSOR_EXTENSIONS[category]:
            continue  # PDF・対象外の拡張子・一時ファイルは取り込まない
        if row["duplicate_of"]:
            row["status"], row["note"] = "unsupported", "duplicate_file"
            stats["duplicate_files"] += 1
            continue

        stats["files"] += 1
        path = repo_root / row["relative_path"]
        try:
            if category == "invoices":
                _import_invoice(path, row, processors[category], year, result)
            else:
                parse = parse_card_csv if category == "credit_card" else parse_bank_csv
                transactions, row_issues = parse(path, processors[category],
                                                 document_key=row["document_key"], target_year=year)
                target = result.card if category == "credit_card" else result.bank
                target.extend(transactions)
                source_type = "card" if category == "credit_card" else "bank"
                for issue in row_issues:
                    result.file_issues.append(_file_issue(
                        source_type, row["document_key"], issue["code"],
                        f"row {issue['row_number']}:{issue['field']}"))
                stats["rows_not_transaction"] += len(row_issues)
        except SheetNotFoundError:
            row["note"] = "template_sheet_not_found"
            stats["template_not_matched"] += 1
            continue
        except TemplateAmbiguousError:
            row["note"] = "template_ambiguous"
            stats["template_not_matched"] += 1
            continue
        except CsvFormatError:
            row["note"] = "format_mismatch"
            stats["failed"] += 1
            result.file_issues.append(_file_issue(
                _source_type(category), row["document_key"], "CSV_HEADER_MISMATCH", None))
            continue
        except Exception:  # noqa: BLE001 - 実データを含む可能性があるため詳細は扱わない
            row["note"] = "read_failed"
            stats["failed"] += 1
            result.file_issues.append(_file_issue(
                _source_type(category), row["document_key"], "FILE_READ_FAILED", None))
            continue
        row["status"], row["note"] = "imported", ""
        stats["imported"] += 1

    _deduplicate(result.card, result.stats["credit_card"])
    _deduplicate(result.bank, result.stats["bank"])
    _flag_duplicate_candidates(result.card, ("use_date", "merchant", "amount"), "amount")
    _flag_duplicate_candidates(
        result.bank, ("transaction_date", "withdrawal_amount", "deposit_amount", "description"),
        "transaction_date")
    for transaction in result.card + result.bank:
        transaction["review_status"] = "needs_review" if transaction["warnings"] else "unreviewed"
    return result


def _source_type(category: str) -> str:
    return {"invoices": "sales", "credit_card": "card", "bank": "bank"}[category]


def _pattern_label(item: dict) -> str:
    """画面・manifest 用の帳票パターン名。template_c は A〜E / unidentified、講師形式は instructor、
    パターンを持たないテンプレート（template_a / template_b）はテンプレート名。"""
    pattern = item["reference"]["template_pattern"]
    if pattern:
        return pattern
    source_name = item["record"]["source_name"]
    return "unidentified" if source_name == template_c.SOURCE_NAME else source_name


class TemplateAmbiguousError(Exception):
    """複数のテンプレートのシートが同じブックにあり、振り分けを決められない。"""


def select_template(path: Path, templates: list):
    """ブックのシート名から、使うテンプレートを 1 つ選ぶ（ファイル名・取引先名・金額は使わない）。

    該当するシートが無ければ SheetNotFoundError、複数あれば TemplateAmbiguousError。
    シートが見つかった後のラベル構造の確認は、各テンプレートの中で行う（合わなければ TEMPLATE_LABEL_MISMATCH）。
    """
    names = set(sheet_names(path))
    candidates = [template for template in templates if template.SHEET_NAME in names]
    if not candidates:
        raise SheetNotFoundError("対応するシートがありません")
    if len(candidates) > 1:
        raise TemplateAmbiguousError("複数のテンプレートに該当します")
    return candidates[0]


def _import_invoice(path: Path, row: dict, templates: list, year: int, result: ProcessResult) -> None:
    template = select_template(path, templates)
    source_hash, extracted, issues = template.read(path)
    record = build_sales_record(extracted=extracted, issues=issues, source_hash=source_hash,
                                source_name=template.SOURCE_NAME)
    record = validate(record, target_year=year)
    result.sales.append({
        "source_key": sales_source_key(source_hash, 1),
        "document_key": row["document_key"],
        "record": record,
        "reference": {key: extracted.get(key) for key in SALES_REFERENCE_FIELDS},
    })


def _deduplicate(transactions: list[dict], stats: Counter) -> None:
    """同じ source_key の取引は最初のものだけを有効にし、以降は duplicate として残す。"""
    first: dict[str, dict] = {}
    for transaction in transactions:
        key = transaction["source_key"]
        if key in first:
            transaction["record_status"] = "duplicate"
            transaction["duplicate_of_document"] = first[key]["document_key"]
            stats["duplicate_transactions"] += 1
        else:
            first[key] = transaction


def _flag_duplicate_candidates(transactions: list[dict], fields: tuple[str, ...],
                               warning_field: str) -> None:
    """有効な取引のうち、主な内容が同じで source_key が異なるものに DUPLICATE_CANDIDATE を付ける（消さない）。"""
    groups: dict[tuple, list[dict]] = defaultdict(list)
    for transaction in transactions:
        if transaction["record_status"] == "active":
            groups[tuple(transaction[f] for f in fields)].append(transaction)
    for members in groups.values():
        if len(members) > 1:
            for transaction in members:
                transaction["warnings"].append(make_issue("DUPLICATE_CANDIDATE", warning_field))


def _file_issue(source_type: str, document_key: str, code: str, field_name: str | None) -> dict:
    return {"source_type": source_type, "source_key": document_key, "document_key": document_key,
            "issue_code": code, "field": field_name}


# ---------------------------------------------------------------- 出力用の行

def _warnings_text(warnings: list[dict]) -> str:
    return ";".join(f"{w['code']}:{w['field']}" for w in warnings)


def sales_rows(result: ProcessResult) -> list[dict]:
    return [
        {
            **record_to_row(item["record"]),
            "source_key": item["source_key"],
            "document_key": item["document_key"],
            "revenue_date": None,  # 売上計上日はプログラムが決めない
            "revenue_date_status": "unconfirmed",
            **item["reference"],
            **item.get("accounting", dict.fromkeys(ACCOUNTING_FIELDS)),
        }
        for item in result.sales
    ]


def decide_accounting_sales_amount(record: dict, reference: dict,
                                   settings: dict | None) -> tuple[int | None, str | None]:
    """会計用の売上金額（税込経理）と、その根拠を返す。決められなければ (None, None)。

    原資料の金額（net / tax / gross 等）は変更しない。逆算・推測はしない。
    - 年度の会計設定が「免税事業者・税込経理」の場合だけ決める（それ以外の組み合わせは未対応のため None）。
    - gross_amount（税込総額）が判明していれば、その値。
    - gross_amount が不明で、帳票上の値が税込の請求総額であると確認済みの場合（template_c のパターン D の
      「ご請求金額」）だけ、その値。
    - template_instructor は、帳票の「小計」（人間の確認により、源泉徴収前の税込報酬額。パーサーは net_amount
      に保持）を使う。最終の「合計」（源泉徴収後）は使わない。条件は _instructor_subtotal を参照。
    - それ以外は None（人間の確認対象）。0 は既知の 0 として扱う。
    """
    if not settings or settings.get("tax_status") != "exempt" \
            or settings.get("consumption_tax_accounting") != "inclusive":
        return None, None
    gross = record.get("gross_amount")
    if _is_amount(gross):
        return gross, "gross_amount"
    if (record.get("source_name") == template_c.SOURCE_NAME
            and reference.get("template_pattern") == "D"):
        total = reference.get("document_total")
        if _is_amount(total):
            return total, "document_total"
    subtotal = _instructor_subtotal(record, reference)
    if subtotal is not None:
        return subtotal, "instructor_subtotal"
    return None, None


def _is_amount(value) -> bool:
    return isinstance(value, int) and not isinstance(value, bool)


def _instructor_subtotal(record: dict, reference: dict) -> int | None:
    """template_instructor の「小計」の値。安全に確認できなければ None。

    net_amount が「小計」ラベルから取得した値であると言えるのは、次をすべて満たす場合だけ:
    - source_name が template_instructor
    - template_pattern が instructor（明細見出しの列・集計ラベルの列と重複なしを、パーサーが確認済み。
      確認できない帳票は instructor_unidentified になり、net_amount は割り当てられない）
    - net_amount が整数で、net_amount に読み取り時の警告（数式キャッシュなし・型違い等）が無い
    """
    if record.get("source_name") != template_instructor.SOURCE_NAME:
        return None
    if reference.get("template_pattern") != template_instructor.PATTERN_NAME:
        return None
    net = record.get("net_amount")
    if not _is_amount(net):
        return None
    if any(w.get("field") == "net_amount" for w in record.get("warnings", [])):
        return None
    return net


def apply_accounting_amounts(result: ProcessResult, settings: dict | None) -> None:
    """各売上に会計用の売上金額を付ける。決められない売上は警告を付け、要確認にする。"""
    for item in result.sales:
        amount, basis = decide_accounting_sales_amount(item["record"], item["reference"], settings)
        item["accounting"] = {"accounting_sales_amount": amount, "accounting_sales_amount_basis": basis}
        if amount is None:
            record = item["record"]
            record["warnings"].append(make_issue("ACCOUNTING_SALES_AMOUNT_UNKNOWN", "accounting_sales_amount"))
            record["review_status"] = "needs_review"


def transaction_rows(transactions: list[dict], columns: list[str]) -> list[dict]:
    return [
        {column: (_warnings_text(t["warnings"]) if column == "warnings" else t.get(column))
         for column in columns}
        for t in transactions
    ]


def unresolved_rows(result: ProcessResult) -> list[dict]:
    """全工程の要確認を 1 つの一覧にする。判断を記入する列は空欄。"""
    items: list[dict] = []

    def add(source_type: str, source_key: str, document_key: str, code: str, field_name):
        items.append({
            "item_id": f"U{len(items) + 1:05d}",
            "source_type": source_type,
            "source_key": source_key,
            "document_key": document_key,
            "issue_code": code,
            "severity": SEVERITY[code],
            "field": field_name,
            "message": MESSAGES[code],
            "decision_status": None,
            "decision_value": None,
            "decision_note": None,
        })

    for issue in result.file_issues:
        add(issue["source_type"], issue["source_key"], issue["document_key"],
            issue["issue_code"], issue["field"])
    for item in result.sales:
        for warning in item["record"]["warnings"]:
            add("sales", item["source_key"], item["document_key"], warning["code"], warning["field"])
        add("sales", item["source_key"], item["document_key"], "REVENUE_DATE_UNCONFIRMED", "revenue_date")
    for source_type, transactions in (("card", result.card), ("bank", result.bank)):
        for transaction in transactions:
            if transaction["record_status"] != "active":
                continue
            for warning in transaction["warnings"]:
                add(source_type, transaction["source_key"], transaction["document_key"],
                    warning["code"], warning["field"])
    return items


def summarize(result: ProcessResult, unresolved: list[dict]) -> dict:
    """manifest と画面表示用の件数（値は含めない）。"""
    def transaction_counts(transactions: list[dict], category: str) -> dict:
        active = [t for t in transactions if t["record_status"] == "active"]
        stats = result.stats.get(category, Counter())
        return {
            "files_imported": stats.get("imported", 0),
            "transactions": len(active),
            "duplicate_transactions": stats.get("duplicate_transactions", 0),
            "rows_not_transaction": stats.get("rows_not_transaction", 0),
            "needs_review": sum(t["review_status"] == "needs_review" for t in active),
        }

    def file_counts(category: str) -> dict:
        stats = result.stats.get(category, Counter())
        return {key: stats.get(key, 0) for key in (
            "files", "imported", "failed", "template_not_matched", "duplicate_files",
            "skipped_by_decision")}

    return {
        "files": {category: file_counts(category) for category in ("invoices", "credit_card", "bank")},
        "sales": {
            "records": len(result.sales),
            "needs_review": sum(i["record"]["review_status"] == "needs_review" for i in result.sales),
            "by_template_pattern": dict(sorted(Counter(
                _pattern_label(i) for i in result.sales).items())),
        },
        "card": transaction_counts(result.card, "credit_card"),
        "bank": transaction_counts(result.bank, "bank"),
        "unresolved": {
            "total": len(unresolved),
            "by_code": dict(sorted(Counter(i["issue_code"] for i in unresolved).items())),
        },
    }

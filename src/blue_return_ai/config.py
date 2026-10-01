"""年間処理の設定ファイル（data/config/annual.json）の読み込み。

設定ファイルは本人が作成し、data/ 配下に置く（Git 管理しない）。記入例は
sample_data/config/annual.example.json を参照。

    {
      "version": 1,
      "sources": {
        "invoices": "data/invoices",
        "bank": "data/bank",
        "credit_card": "data/credit-card",
        "household": "data/household"
      },
      "decisions_dir": "data/decisions",
      "processors": {
        "invoices": {"templates": ["template_c", "template_instructor"]},
        "credit_card": {"format": "card_csv_a"},
        "bank": {"format": "bank_csv_a"}
      }
    }

- invoices は "template": "<名前>"（1 つ）または "templates": [...]（複数）で指定する。
  複数の場合は、各テンプレートのシート名がブックにあるかで振り分ける（ファイル名・取引先名・金額は使わない）。

- パスはリポジトリ直下からの相対パスで書き、data/ 配下だけを指定できる。
- processors に、資料の種類ごとの取込方法を汎用の識別子で指定する（金融機関名・サービス名は書かない）。
  指定の無い種類は取り込まず、資料一覧だけを作る。
- accounting に、年度ごとの会計上の前提（人間が決める）を書く。例:
      "accounting": {"2026": {"tax_status": "exempt", "invoice_registration": false,
                              "consumption_tax_accounting": "inclusive"}}
  tax_status: exempt（免税事業者）/ taxable（課税事業者）、consumption_tax_accounting: inclusive（税込経理）/ exclusive（税抜経理）。
- CSV の書式は、必要なら "options" で上書きできる（encoding / header_row / columns）。
  例: "bank": {"format": "bank_csv_a", "options": {"header_row": 13}}
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path

CONFIG_VERSION = 1
CATEGORIES = ("invoices", "bank", "credit_card", "household")


class ConfigError(Exception):
    """設定ファイルの誤り。メッセージに実データの値を含めない。"""


@dataclass(frozen=True)
class AnnualConfig:
    sha256: str
    sources: dict[str, Path]
    decisions_dir: Path | None
    processors: dict[str, dict]
    accounting: dict[int, dict] = field(default_factory=dict)  # 年度別の会計設定（{年: 設定}）


def _resolve_data_path(value: object, key: str, repo_root: Path) -> Path:
    if not isinstance(value, str) or not value:
        raise ConfigError(f"{key} はリポジトリ直下からの相対パスを文字列で指定してください")
    relative = Path(value)
    if relative.is_absolute() or relative.drive:
        raise ConfigError(f"{key} は相対パスで指定してください")
    resolved = (repo_root / relative).resolve()
    if not resolved.is_relative_to((repo_root / "data").resolve()):
        raise ConfigError(f"{key} は data/ 配下を指定してください")
    return resolved


def load_config(path: Path, repo_root: Path) -> AnnualConfig:
    path = Path(path)
    try:
        raw = path.read_bytes()
        data = json.loads(raw.decode("utf-8-sig"))
    except FileNotFoundError:
        raise ConfigError(
            "設定ファイルが見つかりません（sample_data/config/annual.example.json を参考に作成してください）"
        ) from None
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ConfigError(f"設定ファイルを読み込めません: {type(error).__name__}") from None

    if not isinstance(data, dict) or data.get("version") != CONFIG_VERSION:
        raise ConfigError(f"設定ファイルの version は {CONFIG_VERSION} である必要があります")
    sources = data.get("sources")
    if not isinstance(sources, dict) or not sources:
        raise ConfigError("設定ファイルの sources を指定してください")
    unknown = set(sources) - set(CATEGORIES)
    if unknown:
        raise ConfigError(f"sources に指定できるのは {', '.join(CATEGORIES)} のみです")

    resolved = {
        category: _resolve_data_path(sources[category], f"sources.{category}", repo_root)
        for category in CATEGORIES if category in sources
    }
    decisions = data.get("decisions_dir")
    decisions_dir = (_resolve_data_path(decisions, "decisions_dir", repo_root)
                     if decisions is not None else None)
    return AnnualConfig(
        sha256=hashlib.sha256(raw).hexdigest(),
        sources=resolved,
        decisions_dir=decisions_dir,
        processors=_load_processors(data.get("processors")),
        accounting=_load_accounting(data.get("accounting")),
    )


# 年度別の会計設定で指定できる値
TAX_STATUSES = ("exempt", "taxable")  # 免税事業者 / 課税事業者
CONSUMPTION_TAX_ACCOUNTING = ("inclusive", "exclusive")  # 税込経理 / 税抜経理


def _load_accounting(value: object) -> dict[int, dict]:
    """accounting: {"<年>": {tax_status, invoice_registration, consumption_tax_accounting}} を読む。

    税務上の前提は人間が年度ごとに指定する（プログラムは推測しない）。指定の無い年は空。
    """
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ConfigError("accounting はオブジェクトで指定してください")
    settings: dict[int, dict] = {}
    for year_text, setting in value.items():
        if not (isinstance(year_text, str) and year_text.isdigit() and len(year_text) == 4):
            raise ConfigError("accounting のキーは 4 桁の年（例: \"2026\"）で指定してください")
        label = f"accounting.{year_text}"
        expected = {"tax_status", "invoice_registration", "consumption_tax_accounting"}
        if not isinstance(setting, dict) or set(setting) != expected:
            raise ConfigError(f"{label} には {', '.join(sorted(expected))} をすべて指定してください")
        if setting["tax_status"] not in TAX_STATUSES:
            raise ConfigError(f"{label}.tax_status は {' / '.join(TAX_STATUSES)} のいずれかです")
        if not isinstance(setting["invoice_registration"], bool):
            raise ConfigError(f"{label}.invoice_registration は true / false で指定してください")
        if setting["consumption_tax_accounting"] not in CONSUMPTION_TAX_ACCOUNTING:
            raise ConfigError(
                f"{label}.consumption_tax_accounting は {' / '.join(CONSUMPTION_TAX_ACCOUNTING)} のいずれかです")
        settings[int(year_text)] = dict(setting)
    return settings


PROCESSOR_KEYS = {"invoices": "template", "credit_card": "format", "bank": "format"}


def _load_processors(value: object) -> dict[str, dict]:
    """processors の構造だけを確認する（識別子が実在するかは annual 側で確認する）。"""
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ConfigError("processors はオブジェクトで指定してください")
    unknown = set(value) - set(PROCESSOR_KEYS)
    if unknown:
        raise ConfigError(f"processors に指定できるのは {', '.join(PROCESSOR_KEYS)} のみです")
    processors: dict[str, dict] = {}
    for category, setting in value.items():
        key = PROCESSOR_KEYS[category]
        if category == "invoices" and isinstance(setting, dict) and "templates" in setting:
            processors[category] = _load_invoice_templates(setting)
            continue
        if not isinstance(setting, dict) or not isinstance(setting.get(key), str):
            raise ConfigError(f"processors.{category} には {key} を文字列で指定してください")
        extra = set(setting) - {key, "options"}
        if extra:
            raise ConfigError(f"processors.{category} に指定できるのは {key} と options のみです")
        options = setting.get("options", {})
        if not isinstance(options, dict):
            raise ConfigError(f"processors.{category}.options はオブジェクトで指定してください")
        if category == "invoices" and options:
            raise ConfigError("processors.invoices に options は指定できません")
        processors[category] = {"id": setting[key], "ids": [setting[key]], "options": dict(options)}
    return processors


def _load_invoice_templates(setting: dict) -> dict:
    """invoices の "templates": [...]（複数のテンプレートをシート名で振り分ける）。"""
    if set(setting) != {"templates"}:
        raise ConfigError("processors.invoices には template か templates の一方だけを指定してください")
    templates = setting["templates"]
    if (not isinstance(templates, list) or not templates
            or not all(isinstance(t, str) and t for t in templates)
            or len(set(templates)) != len(templates)):
        raise ConfigError("processors.invoices.templates は重複の無いテンプレート名の配列で指定してください")
    return {"id": ",".join(templates), "ids": list(templates), "options": {}}

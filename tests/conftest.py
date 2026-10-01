"""テスト共通のフィクスチャ。テストには完全な架空データのみを使う。"""

from __future__ import annotations

from pathlib import Path

import pytest

from blue_return_ai import template_a
from blue_return_ai.sales_record import build_sales_record
from blue_return_ai.sample_template_a import write_sample
from blue_return_ai.validation import validate

REPO_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_INVOICE = REPO_ROOT / "sample_data" / "invoices" / "template_a_sample.xlsx"


def process(path: Path, **validate_options) -> dict:
    source_hash, extracted, issues = template_a.read(path)
    record = build_sales_record(
        extracted=extracted,
        issues=issues,
        source_hash=source_hash,
        source_name=template_a.SOURCE_NAME,
    )
    return validate(record, **validate_options)


@pytest.fixture
def sample_invoice() -> Path:
    assert SAMPLE_INVOICE.exists(), "sample_data の架空請求書がありません"
    return SAMPLE_INVOICE


@pytest.fixture
def make_invoice(tmp_path):
    """架空の template_a 請求書を、セルを差し替えて tmp_path に作る。"""
    counter = {"n": 0}

    def _make(overrides: dict | None = None, workbook_hook=None) -> Path:
        counter["n"] += 1
        path = tmp_path / f"invoice_{counter['n']}.xlsx"
        if workbook_hook is None:
            return write_sample(path, overrides)
        from blue_return_ai.sample_template_a import build_workbook
        workbook = build_workbook(overrides)
        workbook_hook(workbook)
        workbook.save(path)
        return path

    return _make


def codes(record: dict) -> set[tuple[str, str | None]]:
    return {(w["code"], w["field"]) for w in record["warnings"]}

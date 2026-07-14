"""예실대비표 구조 스키마 정의."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Literal


CompositionKind = Literal["codes", "remainder", "all"]


@dataclass
class SummaryRule:
    """하단 요약 행 계산 규칙."""

    label: str
    composition: CompositionKind
    codes: list[str] = field(default_factory=list)


@dataclass
class BudgetSchema:
    """파일에서 추론한 표 구조."""

    header_row_count: int
    category_col: int
    code_col: int
    name_col: int
    amount_start_col: int
    amount_headers: list[str]
    amount_col_numbers: list[int]
    subtotal_labels: list[str]
    summary_rules: list[SummaryRule]
    source: str = "unknown"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @property
    def subtotal_normalized(self) -> set[str]:
        return {normalize_label(label) for label in self.subtotal_labels}

    @property
    def summary_normalized(self) -> set[str]:
        return {normalize_label(rule.label) for rule in self.summary_rules}


def normalize_label(value: Any) -> str:
    """공백을 제거해 라벨을 비교·표시 가능한 형태로 만든다."""
    if value is None:
        return ""

    return str(value).replace(" ", "").strip()


def display_label(value: Any) -> str:
    """원본 라벨의 과도한 공백을 제거한 표시용 문자열을 만든다."""
    return normalize_label(value)
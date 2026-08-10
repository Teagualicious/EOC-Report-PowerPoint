"""Pure validation rule engine boundary for Deck Engine."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, Mapping, Sequence


@dataclass(frozen=True)
class Finding:
    severity: str
    code: str
    sheet: str = ""
    row: int | None = None
    column: str = ""
    message: str = ""
    remedy: str = ""

    @property
    def blocking(self) -> bool:
        return self.severity.lower() == "error"


def has_errors(findings: Iterable[Finding]) -> bool:
    return any(finding.blocking for finding in findings)


def validate_ingest(rows: Sequence[Mapping[str, Any]],
                    context: Mapping[str, Any] | None = None) -> list[Finding]:
    """Return ingest-time findings without mutating or writing anything."""
    raise NotImplementedError("Ingest validation rules are implemented in Stage 3")


def validate_fill(payload: Any,
                  context: Mapping[str, Any] | None = None) -> list[Finding]:
    """Return fill-time findings without mutating or writing anything."""
    raise NotImplementedError("Fill validation rules are implemented in Stage 3")

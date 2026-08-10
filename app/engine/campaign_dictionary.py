"""Campaign-name interpretation seam.

Fork Stage 0 fixes the public boundary. Stage 1 implements the v0 identity
passthrough; Stage 8 may replace it with ordered configuration rules without
changing callers.
"""

from __future__ import annotations

from typing import Any


def apply(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    """Return transformed Unified Data rows and human-readable notes."""
    raise NotImplementedError("Campaign dictionary v0 is implemented in Stage 1")

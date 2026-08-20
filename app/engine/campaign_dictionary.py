"""Campaign-name interpretation seam.

Fork Stage 0 fixes the public boundary. Stage 1 implements the v0 identity
passthrough; Stage 8 may replace it with ordered configuration rules without
changing callers.
"""

from __future__ import annotations

from typing import Any


def apply(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    """Return identity-preserved rows and deterministic analyst notes.

    Stage 1 intentionally has no campaign aliases or normalization rules.  The
    explicit identity pass keeps the seam stable for Stage 8 while making the
    absence of interpretation visible to the analyst and to reconciliation
    reports.  Input dictionaries are copied so a later dictionary version
    cannot mutate parser-owned data in place.
    """
    preserved = [dict(row) for row in rows]
    campaigns = sorted({
        str(row.get("campaign_name", ""))
        for row in preserved
        if row.get("campaign_name", "")
    })
    notes = [
        "Campaign dictionary v0 identity passthrough: no campaign rules were applied.",
        f"Campaign dictionary v0 retained {len(campaigns)} campaign name(s) unchanged.",
    ]
    return preserved, notes

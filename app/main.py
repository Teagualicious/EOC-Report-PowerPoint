"""Deck Engine desktop entry point.

Fork Stage 0 intentionally has no analyst UI. The real one-window application
lands in Stage 5; until then the CLI is the only supported shell.
"""

from __future__ import annotations

import logging
import sys

from config.logging_setup import setup_logging
from config.paths import ensure_dirs


def main() -> int:
    ensure_dirs()
    setup_logging()
    logging.getLogger(__name__).info("Deck Engine Stage 0 placeholder started")
    print(
        "Deck Engine's analyst desktop application is not available in Fork "
        "Stage 0. Use: python -m app.cli list-templates",
        file=sys.stderr,
    )
    return 2


if __name__ == "__main__":
    raise SystemExit(main())

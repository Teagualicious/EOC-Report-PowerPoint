"""Application-wide logging configuration.

Every module obtains its logger with ``logging.getLogger(__name__)`` and
never attaches handlers itself; this module owns the handlers. Logs land
in ``workspace/logs/ingestion_engine.log`` (rotating, 3 backups) so a support
person can ask an AE to send the file when something goes wrong on a
machine we can't touch.
"""

import logging
import logging.handlers
import os

from config.paths import LOGS_DIR

LOG_FILE = os.path.join(LOGS_DIR, "ingestion_engine.log")
_FORMAT = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"


def setup_logging(level=logging.INFO):
    """Configure the root logger. Safe to call more than once."""
    root = logging.getLogger()
    if root.handlers:  # already configured (e.g. tests or re-entry)
        return root

    root.setLevel(level)

    os.makedirs(LOGS_DIR, exist_ok=True)
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    file_handler.setFormatter(logging.Formatter(_FORMAT))
    root.addHandler(file_handler)

    # Console shows warnings and errors so launcher users see real problems
    console = logging.StreamHandler()
    console.setLevel(logging.WARNING)
    console.setFormatter(logging.Formatter(_FORMAT))
    root.addHandler(console)

    return root

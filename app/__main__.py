"""Allow ``python -m app`` as the manual launch path."""

from app.main import main

raise SystemExit(main())

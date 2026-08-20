"""Print the ICAI Division I Schedule III hierarchy extracted from the Guidance Note."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.mapping.schedule_iii_mapper import main


if __name__ == "__main__":
    raise SystemExit(main())

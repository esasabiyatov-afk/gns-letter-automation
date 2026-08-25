from __future__ import annotations

import os


# Office test builds must exercise the real read-only Tolubay adapter. Secrets
# are still requested by the application and are never embedded in the EXE.
os.environ.setdefault("GNS_ABS_MODE", "tolubay")
os.environ.setdefault("GNS_PROCESSING_WORKERS", "1")

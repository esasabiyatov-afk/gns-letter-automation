from __future__ import annotations

import os


# Portable office edition: Outlook sends to the recipient configured for the
# letter.  Test-recipient routing is intentionally not enabled here.
os.environ.setdefault("GNS_ABS_MODE", "tolubay")
os.environ.setdefault("GNS_TOLUBAY_VERIFY_TLS", "false")
os.environ.setdefault("GNS_OUTLOOK_ALLOW_INSECURE_CERTIFICATE", "true")
os.environ.setdefault("GNS_PROCESSING_WORKERS", "1")

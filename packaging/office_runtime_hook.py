from __future__ import annotations

import os


# Office builds use the real read-only Tolubay adapter. Secrets are still
# requested by the application and are never embedded in the EXE.
os.environ.setdefault("GNS_ABS_MODE", "tolubay")
# The bank confirmed that the currently deployed Tolubay certificate is
# expired.  Settings still restrict this exception to https://ob.tolubay.kg.
os.environ.setdefault("GNS_TOLUBAY_VERIFY_TLS", "false")
# During Outlook synchronization, confirm only the known Outlook certificate
# dialog.  Other dialogs and all send actions remain under user control.
os.environ.setdefault("GNS_OUTLOOK_ALLOW_INSECURE_CERTIFICATE", "true")
# Office test builds send only to the hard-coded approved Gmail recipient.
os.environ.setdefault("GNS_OUTLOOK_ALLOW_TEST_SEND", "true")
os.environ.setdefault("GNS_PROCESSING_WORKERS", "1")

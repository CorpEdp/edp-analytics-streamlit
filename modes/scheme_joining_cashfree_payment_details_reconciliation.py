from core.config import LEGACY_DIR
from core.loader import execute_legacy_script

LEGACY_SOURCE = LEGACY_DIR / "ERP-scheme_joining_cashfree_payment_details_reconciliation.py"


def render() -> None:
    execute_legacy_script(LEGACY_SOURCE)

from core.config import LEGACY_DIR
from core.loader import execute_legacy_script

LEGACY_SOURCE = LEGACY_DIR / "ERP-CashfreeReconciliation.py"


def render() -> None:
    execute_legacy_script(LEGACY_SOURCE)

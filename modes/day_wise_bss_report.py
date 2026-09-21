from core.config import LEGACY_DIR
from core.loader import execute_legacy_script

LEGACY_SOURCE = LEGACY_DIR / "ERP-Bss_Day_wise_Report.py"


def render() -> None:
    execute_legacy_script(LEGACY_SOURCE)

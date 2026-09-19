from core.config import LEGACY_DIR
from core.loader import execute_legacy_script

LEGACY_SOURCE = LEGACY_DIR / "App-Employee_referral_Performance_Report.py"


def render() -> None:
    execute_legacy_script(LEGACY_SOURCE)

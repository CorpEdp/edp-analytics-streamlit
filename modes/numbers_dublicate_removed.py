from core.config import LEGACY_DIR
from core.loader import execute_legacy_script

LEGACY_SOURCE = LEGACY_DIR / "CustomNumbersDublicateRemoved.py"


def render() -> None:
    execute_legacy_script(LEGACY_SOURCE)

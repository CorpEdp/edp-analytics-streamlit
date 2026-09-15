#!/usr/bin/env python3
"""Generate a safe registry and mode summary for all legacy Streamlit modules."""

from __future__ import annotations

import ast
import hashlib
import json
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent.parent
MODULES_DIR = BASE_DIR / "modules"
MODES_DIR = BASE_DIR / "modes"
REGISTRY_PATH = MODES_DIR / "registry.py"


def safe_name(name: str) -> str:
    cleaned = name
    cleaned = cleaned.replace("-", "_")
    cleaned = cleaned.replace("&", "_")
    cleaned = cleaned.replace(" ", "_")
    cleaned = cleaned.replace(".", "_")
    cleaned = "".join(ch if ch.isalnum() or ch == "_" else "_" for ch in cleaned)
    cleaned = cleaned.strip("_")
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    return cleaned.lower()


def module_name_for_legacy(path: Path) -> str:
    stem = path.stem
    if stem.lower().startswith("app"):
        stem = stem[3:]
    elif stem.lower().startswith("erp"):
        stem = stem[3:]
    return safe_name(stem)


def read_source(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def get_imports(source: str) -> list[str]:
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return []
    imports: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                imports.append(alias.name)
        elif isinstance(node, ast.ImportFrom):
            if node.module:
                imports.append(node.module)
    return sorted(set(imports))


def discover_modules() -> list[dict[str, Any]]:
    discovered: list[dict[str, Any]] = []
    for path in sorted(MODULES_DIR.glob("*.py")):
        if path.name == "__init__.py":
            continue
        source = read_source(path)
        discovered.append(
            {
                "original_file": path.name,
                "module_name": module_name_for_legacy(path),
                "sha256": hashlib.sha256(source.encode("utf-8")).hexdigest(),
                "imports": get_imports(source),
                "size_bytes": path.stat().st_size,
            }
        )
    return discovered


def generate_registry(modules: list[dict[str, Any]]) -> str:
    lines = [
        '"""Auto-generated registry for the unified Streamlit app."""',
        "",
        "MODE_REGISTRY = [",
        "    {",
        '        "id": "home",',
        '        "name": "Dashboard / Home",',
        '        "description": "Unified EDP Analytics home screen",',
        '        "category": "General",',
        '        "module": "modes.home",',
        '        "source": "app.py",',
        "    },",
    ]

    for item in modules:
        safe_id = item["module_name"]
        label = item["original_file"]
        label = label.replace(".py", "")
        label = label.replace("App-", "")
        label = label.replace("ERP-", "")
        label = label.replace("_", " ")
        label = " ".join(part for part in label.split() if part)
        label = label.title()
        lines.extend([
            "    {",
            f'        "id": "{safe_id}",',
            f'        "name": "{label}",',
            f'        "description": "Legacy Streamlit mode converted from {item["original_file"]}",',
            '        "category": "Legacy",',
            f'        "module": "modes.{safe_id}",',
            f'        "source": "modules/{item["original_file"]}",',
            "    },",
        ])

    lines.append("]")
    return "\n".join(lines) + "\n"


def main() -> None:
    modules = discover_modules()
    registry_text = generate_registry(modules)

    REGISTRY_PATH.write_text(registry_text, encoding="utf-8")

    report = {
        "discovered_files": [m["original_file"] for m in modules],
        "converted_files": [m["module_name"] for m in modules],
        "skipped_files": [],
        "failed_files": [],
        "import_problems": [],
        "duplicate_files": [],
    }
    (BASE_DIR / "mode_registry_report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")

    print(f"Discovered {len(modules)} module(s)")
    print(f"Updated registry: {REGISTRY_PATH}")


if __name__ == "__main__":
    main()

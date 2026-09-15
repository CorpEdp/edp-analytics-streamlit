"""Load one legacy Streamlit script only when its mode is selected."""

import ast
from pathlib import Path
from types import ModuleType
from typing import Any


class _PageConfigRemover(ast.NodeTransformer):
    def visit_Expr(self, node: ast.Expr) -> Any:
        value = node.value
        if (
            isinstance(value, ast.Call)
            and isinstance(value.func, ast.Attribute)
            and value.func.attr == "set_page_config"
        ):
            return ast.copy_location(ast.Pass(), node)
        return self.generic_visit(node)


def execute_legacy_script(source_path: Path) -> dict[str, Any]:
    """Execute a legacy app in an isolated namespace at render time."""
    source = source_path.read_text(encoding="utf-8", errors="replace")
    tree = ast.parse(source, filename=str(source_path))
    tree = _PageConfigRemover().visit(tree)
    ast.fix_missing_locations(tree)

    namespace: dict[str, Any] = {
        "__name__": "__main__",
        "__file__": str(source_path),
        "__package__": None,
    }
    exec(compile(tree, str(source_path), "exec"), namespace, namespace)
    return namespace


def load_selected_mode(module_name: str) -> ModuleType:
    """Import only the selected generated mode module."""
    import importlib

    return importlib.import_module(module_name)

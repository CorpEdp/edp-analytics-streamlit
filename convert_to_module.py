"""
convert_to_module.py

Automatically converts an old standalone Streamlit app.py (code running at
top level) into a proper module for the merged project:
  - Wraps the whole script body inside def run():
  - Keeps top-level imports and top-level `def`/`class` blocks OUTSIDE run()
    (functions/classes should stay defined at module level; only the
    executing statements get indented into run())
  - Removes any st.set_page_config(...) call (only main.py may call it)
  - Adds a LABEL = "..." line
  - Writes the result as a new file with a clean, valid module name

USAGE (run from anywhere, e.g. inside your D:\EDP\Project folder):

    python convert_to_module.py "App-BranchEmployeeWiseReferralReport.py"

    # convert every App-*.py / ERP-*.py file in the current folder at once:
    python convert_to_module.py --all

Output goes into a "modules_converted" folder next to the script. Review
each output file, then move it into your real modules/ folder.

IMPORTANT — read this before trusting the output blindly:
  This is a best-effort AUTOMATED indent wrapper. It handles the common
  case (a script that is: imports, then a long flat block of st.* calls,
  maybe with some already-existing function defs). For files with unusual
  structure (e.g. top-level code inside if __name__ == "__main__": blocks,
  or heavy use of global state) you should still spot-check the output.
"""

import argparse
import ast
import glob
import os
import re
import sys


def make_module_name(filename: str) -> str:
    """Turn 'App-BranchEmployeeWiseReferralReport.py' into
    'branch_employee_wise_referral_report.py' (valid Python module name)."""
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(r"^(App|ERP)[-_]?", "", name, flags=re.IGNORECASE)
    # split CamelCase -> words
    name = re.sub(r"(?<!^)(?=[A-Z])", "_", name)
    # normalize any separator (-, &, space, multiple underscores) to single _
    name = re.sub(r"[\s\-&]+", "_", name)
    name = re.sub(r"_+", "_", name)
    name = name.strip("_").lower()
    return name + ".py"


def make_label(filename: str) -> str:
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(r"^(App|ERP)[-_]?", "", name, flags=re.IGNORECASE)
    name = re.sub(r"(?<!^)(?=[A-Z])", " ", name)
    name = re.sub(r"[_\-]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def strip_page_config(source: str) -> str:
    """Remove any st.set_page_config(...) call, including multi-line ones."""
    tree = ast.parse(source)
    lines = source.splitlines(keepends=True)
    to_remove = set()

    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            func = node.func
            is_page_config = (
                isinstance(func, ast.Attribute)
                and func.attr == "set_page_config"
            )
            if is_page_config:
                start = node.lineno
                end = getattr(node, "end_lineno", node.lineno)
                for ln in range(start, end + 1):
                    to_remove.add(ln)

    kept = [line for i, line in enumerate(lines, start=1) if i not in to_remove]
    return "".join(kept)


def wrap_in_run(source: str) -> str:
    """
    Split the module into:
      - top-level import statements  -> stay at module level
      - top-level function/class defs -> stay at module level
      - everything else at top level  -> gets indented into run()
    Order in the file is preserved for the run()-body statements.
    """
    tree = ast.parse(source)
    lines = source.splitlines(keepends=True)

    header_chunks = []   # imports, defs, classes -> stay top-level
    body_chunks = []      # everything else -> goes inside run()

    for node in tree.body:
        start = node.lineno - 1
        end = getattr(node, "end_lineno", node.lineno)
        chunk = "".join(lines[start:end])

        if isinstance(node, (ast.Import, ast.ImportFrom, ast.FunctionDef,
                              ast.AsyncFunctionDef, ast.ClassDef)):
            header_chunks.append(chunk)
        else:
            body_chunks.append(chunk)

    header = "".join(header_chunks).rstrip() + "\n"

    body_source = "".join(body_chunks)
    indented_lines = []
    for line in body_source.splitlines():
        indented_lines.append(("    " + line) if line.strip() else "")
    indented_body = "\n".join(indented_lines)

    return header, indented_body


def convert_file(path: str, out_dir: str):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        source = f.read()

    try:
        source = strip_page_config(source)
        header, indented_body = wrap_in_run(source)
    except SyntaxError as e:
        print(f"  [SKIP] {path}: could not parse as valid Python ({e}). "
              f"Fix syntax errors first, or convert this one by hand.")
        return None

    label = make_label(path)
    out_name = make_module_name(path)
    out_path = os.path.join(out_dir, out_name)

    result = f'''"""
Auto-converted from: {os.path.basename(path)}
Review this file before using — the converter does a best-effort wrap;
double-check indentation around any unusual control flow (loops, if/else
blocks that span large sections, etc.).
"""

{header}
LABEL = "{label}"


def run():
{indented_body if indented_body.strip() else "    pass"}
'''

    os.makedirs(out_dir, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(result)

    print(f"  [OK] {path} -> {out_path}")
    return out_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("files", nargs="*", help="Specific .py files to convert")
    parser.add_argument("--all", action="store_true",
                         help="Convert every App-*.py and ERP-*.py file in the current folder")
    parser.add_argument("--out", default="modules_converted",
                         help="Output folder (default: modules_converted)")
    args = parser.parse_args()

    targets = list(args.files)
    if args.all:
        targets += glob.glob("App-*.py") + glob.glob("App*.py") + glob.glob("ERP-*.py") + glob.glob("ERP*.py")
        targets += glob.glob("*.py")
        targets = sorted(set(targets))
        targets = [t for t in targets if os.path.basename(t) not in
                   ("main.py", "convert_to_module.py") and not t.startswith("modules")]

    if not targets:
        print("No files given. Use: python convert_to_module.py \"File.py\"  or  --all")
        sys.exit(1)

    print(f"Converting {len(targets)} file(s) -> ./{args.out}/\n")
    ok, failed = 0, 0
    for path in targets:
        if not os.path.exists(path):
            print(f"  [MISSING] {path}")
            failed += 1
            continue
        result = convert_file(path, args.out)
        ok += 1 if result else 0
        failed += 0 if result else 1

    print(f"\nDone. {ok} converted, {failed} skipped/failed.")
    print(f"Now review the files in ./{args.out}/ then copy them into your modules/ folder.")


if __name__ == "__main__":
    main()
# tests/check_docs.py
"""Check that folder READMEs, module docstrings and public-function docstrings exist.

Portable, stdlib-only. Copy this file to another repo and edit the CONFIG block.
Rules enforced (see AGENTS.md, "Docstring convention" / "Folder README convention"):

1. README  - every folder under CODE_ROOTS that directly contains tracked .py files
             has a README.md with REQUIRED_README_SECTIONS, and its "## Scripts" table
             names every tracked .py of that folder. Folders in COVERED_BY_PARENT are
             documented by the mapped parent's README instead (same table rule).
             Every folder in README_ROOTS has a README.md (no section check).
2. MODULE  - every tracked .py (except EXEMPT_FILES / __init__.py) has a module
             docstring whose first non-empty line is a real summary, not an
             authorship header (Created on / @author / Updated on).
3. FUNCS   - every public (no leading "_") module-level def/class and public method:
             docstring REQUIRED if the module is a library (LIBRARY_DIRS or imported
             by >= LIBRARY_IMPORTERS other tracked modules); otherwise REQUIRED only
             if the body spans more than MIN_BODY_LINES lines. Nested functions and
             EXEMPT_FUNCS inside FIG_GLOB files are ignored.

Usage:
    python tests/check_docs.py                     # full check; exit 1 on findings
    python tests/check_docs.py --summary           # coverage per folder; exit 0
    python tests/check_docs.py --paths ostram/pipeline tests
    python tests/check_docs.py --compare-ast main  # docstring-only diff vs a git ref
    python tests/check_docs.py --selftest          # run the built-in fixture tests

Notes:
    - Outside a git checkout, untracked scratch files are included (tracked_py_files falls back
      to rglob when `git ls-files` is unavailable).
"""
from __future__ import annotations

import argparse
import ast
import fnmatch
import subprocess
import sys
import tempfile
import warnings
from collections import defaultdict
from pathlib import Path

warnings.filterwarnings("ignore", category=SyntaxWarning)   # ast.parse of legacy files is noisy

# ----------------------------------------------------------------------------- CONFIG
# Edit this block when copying the checker to another repo. Paths are POSIX,
# relative to the repo root (the folder that contains AGENTS.md).
CODE_ROOTS = ["ostram", "tests", "timeslice_module/scripts"]
README_ROOTS = ["inputs", "config", "model"]          # workspace/ is gitignored and has no README
EXEMPT_DIRS = {"__pycache__"}
COVERED_BY_PARENT = {          # folder -> folder whose README.md documents it (its Scripts table must list them)
    "ostram/pipeline": "ostram",      # one module (orchestration.py): listed as pipeline/orchestration.py
    "ostram/validation": "ostram",    # one module (profile.py): listed as validation/profile.py
}
LIBRARY_DIRS = [                # none: the shared top-level modules (paths, profiles, terminal,
]                               # profile_workspace, examples) already qualify through LIBRARY_IMPORTERS
LIBRARY_IMPORTERS = 2          # imported by >= N other modules => treated as library
EXEMPT_FILES = {
    "docs/conf.py",            # Sphinx configuration, outside CODE_ROOTS; kept explicit
}
EXEMPT_FUNCS = {"main"}        # only inside files matching FIG_GLOB
FIG_GLOB = "figures/*/fig_*.py"   # OSTRAM has no figure scripts; this pattern matches no tracked file
MIN_BODY_LINES = 30
REQUIRED_README_SECTIONS = ["## Scripts", "## How to run", "## Gotchas"]
AUTHORSHIP_PREFIXES = ("created on", "@author", "updated on", "-*-")
# --------------------------------------------------------------------------- /CONFIG


def repo_root() -> Path:
    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        if (parent / "AGENTS.md").exists() or (parent / ".git").exists():
            return parent
    return here.parents[2]


def posix(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def tracked_py_files(root: Path, roots: list[str]) -> list[Path]:
    """Tracked .py files under roots (git ls-files), or all .py files when not a git repo."""
    files: list[Path] = []
    try:
        out = subprocess.run(
            ["git", "-C", str(root), "ls-files", "-z", "--", *roots],
            check=True, capture_output=True, text=True, encoding="utf-8",
        ).stdout
        # -z / NUL-separated so non-ASCII paths are not quoted (unlike the default newline output).
        files = [root / p for p in out.split("\0") if p.endswith(".py")]
    except (subprocess.CalledProcessError, FileNotFoundError):
        for r in roots:
            files.extend(p for p in (root / r).rglob("*.py"))
    files = [f for f in files if not (set(f.parts) & EXEMPT_DIRS)]
    files = [f for f in files if posix(f, root) not in EXEMPT_FILES]
    return sorted(set(files))


def parse(path: Path) -> ast.Module | None:
    try:
        return ast.parse(path.read_text(encoding="utf-8-sig", errors="replace"), filename=str(path))
    except SyntaxError:
        return None


# ------------------------------------------------------------------ 1. README checks
def folders_needing_readme(files: list[Path], root: Path) -> dict[Path, list[Path]]:
    by_dir: dict[Path, list[Path]] = defaultdict(list)
    for f in files:
        if f.name == "__init__.py":
            continue
        d = f.parent
        rel = posix(d, root)
        if rel in COVERED_BY_PARENT:
            d = root / COVERED_BY_PARENT[rel]      # documented by the parent's README
        by_dir[d].append(f)
    return by_dir


def scripts_section(text: str) -> str:
    """Return the body of '## Scripts' up to the next '## ' heading ('' if absent)."""
    start = text.find("## Scripts")
    if start < 0:
        return ""
    rest = text[start + len("## Scripts"):]
    nxt = rest.find("\n## ")
    return rest if nxt < 0 else rest[:nxt]


def scripts_table_names(table: str) -> set[str]:
    """Return the first-column script name/path of each Markdown table row in `table`.

    Only the first `| ... |` cell of each row counts as naming a script (e.g. `` `x.py` `` or
    `` `sub/x.py` ``); mentions of a filename elsewhere in the table (another column, prose) do
    not. Header and separator rows (`---`) are skipped.
    """
    names: set[str] = set()
    for line in table.splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        cell = line[1:].split("|", 1)[0].strip().strip("`").strip()
        if not cell or set(cell) <= {"-", ":"}:
            continue
        names.add(cell)
    return names


def check_readmes(by_dir: dict[Path, list[Path]], root: Path, paths_filter: list[Path]) -> list[str]:
    errors: list[str] = []
    for d, pys in sorted(by_dir.items()):
        if paths_filter and not any(d == p or p in d.parents for p in paths_filter):
            continue
        readme = d / "README.md"
        rel = posix(d, root)
        if not readme.exists():
            errors.append(f"{rel}/README.md: missing")
            continue
        text = readme.read_text(encoding="utf-8", errors="replace")
        for sec in REQUIRED_README_SECTIONS:
            if sec not in text:
                errors.append(f"{rel}/README.md: missing section '{sec}'")
        names = scripts_table_names(scripts_section(text))
        for py in pys:
            if not any(n == py.name or n.endswith("/" + py.name) for n in names):
                errors.append(f"{rel}/README.md: '## Scripts' does not list {py.name}")
    for r in README_ROOTS:
        d = root / r
        if paths_filter and not any(d == p or p in d.parents or d in p.parents for p in paths_filter):
            continue
        if d.is_dir() and not (d / "README.md").exists():
            errors.append(f"{r}/README.md: missing")
    return errors


# ------------------------------------------------------------ 2./3. docstring checks
def first_doc_line(node: ast.AST) -> str:
    doc = ast.get_docstring(node, clean=True) or ""
    for line in doc.splitlines():
        if line.strip():
            return line.strip()
    return ""


def module_docstring_ok(tree: ast.Module) -> bool:
    line = first_doc_line(tree).lower()
    return bool(line) and not line.startswith(AUTHORSHIP_PREFIXES)


def import_counts(files: list[Path]) -> dict[str, int]:
    """stem -> number of OTHER tracked modules that import it (by bare module name)."""
    stems = {f.stem: f for f in files}
    importers: dict[str, set[Path]] = defaultdict(set)
    for f in files:
        tree = parse(f)
        if tree is None:
            continue
        for node in ast.walk(tree):
            names: list[str] = []
            if isinstance(node, ast.Import):
                names = [a.name.split(".")[-1] for a in node.names]
            elif isinstance(node, ast.ImportFrom):
                names = ([node.module.split(".")[-1]] if node.module else []) + [a.name for a in node.names]
            for n in names:
                if n in stems and stems[n] != f:
                    importers[n].add(f)
    return {k: len(v) for k, v in importers.items()}


def is_library(path: Path, root: Path, counts: dict[str, int]) -> bool:
    rel = posix(path.parent, root)
    if any(rel == d or rel.startswith(d + "/") for d in LIBRARY_DIRS):
        return True
    return counts.get(path.stem, 0) >= LIBRARY_IMPORTERS


def public_defs(tree: ast.Module):
    """Yield (node, qualname) for module-level public defs/classes and public methods."""
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            if node.name.startswith("_"):
                continue
            yield node, node.name
            if isinstance(node, ast.ClassDef):
                for sub in node.body:
                    if isinstance(sub, (ast.FunctionDef, ast.AsyncFunctionDef)) and not sub.name.startswith("_"):
                        yield sub, f"{node.name}.{sub.name}"


def body_lines(node: ast.AST) -> int:
    return (getattr(node, "end_lineno", node.lineno) or node.lineno) - node.lineno + 1


def check_module(path: Path, root: Path, counts: dict[str, int]) -> tuple[list[str], int, int]:
    """Return (errors, n_required_defs, n_documented_required_defs) for one file."""
    rel = posix(path, root)
    errors: list[str] = []
    if path.name == "__init__.py":
        return errors, 0, 0
    tree = parse(path)
    if tree is None:
        return [f"{rel}: cannot parse (SyntaxError)"], 0, 0
    if not module_docstring_ok(tree):
        errors.append(f"{rel}: module docstring missing or authorship-only")
    lib = is_library(path, root, counts)
    is_fig = fnmatch.fnmatch(rel, FIG_GLOB)
    required = documented = 0
    for node, qual in public_defs(tree):
        if is_fig and qual in EXEMPT_FUNCS:
            continue
        if not (lib or body_lines(node) > MIN_BODY_LINES):
            continue
        required += 1
        if ast.get_docstring(node):
            documented += 1
        else:
            why = "library module" if lib else f">{MIN_BODY_LINES} lines"
            errors.append(f"{rel}:{node.lineno}: {qual} has no docstring ({why})")
    return errors, required, documented


# -------------------------------------------------------------------- --compare-ast
def strip_docstrings(tree: ast.AST) -> ast.AST:
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            body = node.body
            if body and isinstance(body[0], ast.Expr) and isinstance(getattr(body[0], "value", None), ast.Constant) \
                    and isinstance(body[0].value.value, str):
                del body[0]
                if not body:
                    body.append(ast.Pass())
    return tree


def compare_ast(root: Path, ref: str, roots: list[str]) -> list[str]:
    """Files changed vs `ref` whose logic (AST minus docstrings) differs. New files are skipped."""
    out = subprocess.run(
        ["git", "-C", str(root), "diff", "--name-only", ref, "--", *roots],
        check=True, capture_output=True, text=True, encoding="utf-8",
    ).stdout
    findings: list[str] = []
    for rel in [l for l in out.splitlines() if l.endswith(".py")]:
        work = root / rel
        if not work.exists():
            findings.append(f"{rel}: deleted in working tree")
            continue
        show = subprocess.run(
            ["git", "-C", str(root), "show", f"{ref}:{rel}"],
            capture_output=True, text=True, encoding="utf-8",
        )
        if show.returncode != 0:
            print(f"note: {rel} is new (not in {ref}); skipped")
            continue
        try:
            a = ast.dump(strip_docstrings(ast.parse(show.stdout)))
            work_text = work.read_text(encoding="utf-8-sig", errors="replace")
            b = ast.dump(strip_docstrings(ast.parse(work_text)))
        except SyntaxError as exc:
            findings.append(f"{rel}: SyntaxError ({exc})")
            continue
        if a != b:
            findings.append(f"{rel}: logic changed (AST differs after stripping docstrings)")
    return findings


# --------------------------------------------------------------------------- driver
def run(root: Path, paths: list[str], summary: bool) -> int:
    """Run the README/module/function checks and print the findings (or a coverage summary).

    Args:
        root: repo root (as returned by repo_root()).
        paths: repo-relative paths to limit the check to; empty means everything.
        summary: if True, print per-folder coverage counts and return 0 instead of findings.

    Returns:
        0 if there are no findings (or summary was requested), 1 otherwise.
    """
    files = tracked_py_files(root, CODE_ROOTS)
    filt = [root / p for p in paths]
    counts = import_counts(files)
    by_dir = folders_needing_readme(files, root)
    errors = check_readmes(by_dir, root, filt)
    per_dir: dict[str, list[int]] = defaultdict(lambda: [0, 0, 0, 0])  # files, mod_ok, req, doc
    for f in files:
        if filt and not any(f == p or p in f.parents for p in filt):
            continue
        errs, req, doc = check_module(f, root, counts)
        errors.extend(errs)
        if f.name == "__init__.py":
            continue
        d = per_dir[posix(f.parent, root)]
        d[0] += 1
        d[1] += 0 if any(e.endswith("module docstring missing or authorship-only") for e in errs) else 1
        d[2] += req
        d[3] += doc
    if summary:
        print(f"{'folder':48} files  mod_ok  funcs_req  funcs_doc")
        for k in sorted(per_dir):
            v = per_dir[k]
            print(f"{k:48} {v[0]:5} {v[1]:7} {v[2]:10} {v[3]:10}")
        tot = [sum(v[i] for v in per_dir.values()) for i in range(4)]
        print(f"{'TOTAL':48} {tot[0]:5} {tot[1]:7} {tot[2]:10} {tot[3]:10}")
        print(f"README findings: {sum(1 for e in errors if 'README.md' in e)}")
        return 0
    for e in errors:
        print(e)
    n_files = sum(v[0] for v in per_dir.values())
    print(f"{n_files} files checked, {len(errors)} missing")
    return 1 if errors else 0


# -------------------------------------------------------------------------- selftest
def _write(p: Path, text: str) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8")


def selftest() -> int:
    """Fixture-based tests of every rule. Prints one line per case; exit 1 on failure."""
    global CODE_ROOTS, README_ROOTS, COVERED_BY_PARENT, LIBRARY_DIRS, EXEMPT_FILES
    saved = (CODE_ROOTS, README_ROOTS, COVERED_BY_PARENT, LIBRARY_DIRS, EXEMPT_FILES)
    failures = 0

    def case(name: str, ok: bool, detail: str = "") -> None:
        nonlocal failures
        print(f"[{'ok' if ok else 'FAIL'}] {name}" + (f"  {detail}" if detail and not ok else ""))
        failures += 0 if ok else 1

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        CODE_ROOTS, README_ROOTS, COVERED_BY_PARENT, LIBRARY_DIRS, EXEMPT_FILES = (
            ["scripts"], ["inputs"], {"scripts/pkg/sub": "scripts/pkg"}, ["scripts/lib"], {"scripts/pkg/ignored.py"})
        (root / "AGENTS.md").write_text("x", encoding="utf-8")
        (root / "inputs").mkdir()                            # README_ROOTS folder without README.md
        # pkg: README lists a.py and sub/c.py but not b.py; b.py has authorship-only docstring;
        # long undocumented def in a.py
        _write(root / "scripts/pkg/a.py", '"""Real summary."""\n\ndef long_fn():\n' + "    x = 1\n" * 35 + "\n\ndef short_fn():\n    return 1\n")
        _write(root / "scripts/pkg/b.py", '"""Created on 2021\n@author: someone\n"""\nX = 1\n')
        _write(root / "scripts/pkg/ignored.py", "no docstring at all\n")
        # a.py's row mentions "b.py" in prose (a different column) on purpose: the first-column
        # match must still flag b.py as unlisted rather than matching that prose mention.
        _write(root / "scripts/pkg/README.md",
               "# scripts/pkg\n\n## Scripts\n\n| Script | Notes |\n|---|---|\n"
               "| `a.py` | wraps b.py |\n| `sub/c.py` | - |\n\n## How to run\n\n## Gotchas\n")
        _write(root / "scripts/pkg/sub/c.py", '"""C."""\nX = 1\n')   # covered by parent README (listed there)
        # lib: every public def needs a docstring; private and nested ignored
        _write(root / "scripts/lib/util.py", '"""Lib."""\n\ndef pub():\n    def inner():\n        pass\n    return 1\n\ndef _priv():\n    return 2\n\nclass K:\n    def m(self):\n        return 3\n    def _h(self):\n        return 4\n')
        _write(root / "scripts/lib/README.md", "## Scripts\n`util.py`\n## How to run\n## Gotchas\n")
        # nolib: short undocumented def is fine; imported-by-2 rule
        _write(root / "scripts/nolib/x.py", '"""X."""\nimport shared\ndef f():\n    return 1\n')
        _write(root / "scripts/nolib/y.py", '"""Y."""\nfrom shared import g\n')
        _write(root / "scripts/nolib/shared.py", '"""Shared."""\ndef g():\n    return 1\n')
        _write(root / "scripts/nolib/README.md", "## Scripts\n`x.py` `y.py` `shared.py`\n## How to run\n## Gotchas\n")
        # inputs/ has no README

        files = tracked_py_files(root, CODE_ROOTS)
        case("untracked_ignored (EXEMPT_FILES respected)", all(f.name != "ignored.py" for f in files))
        counts = import_counts(files)
        by_dir = folders_needing_readme(files, root)
        r_errs = check_readmes(by_dir, root, [])
        case("readme_missing_script_row", any("does not list b.py" in e for e in r_errs), str(r_errs))
        case("readme_first_column_match_ignores_prose_mention",
             any("does not list b.py" in e for e in r_errs), str(r_errs))
        case("readme_covered_by_parent_listed_in_parent", not any("pkg/sub" in e or "c.py" in e for e in r_errs), str(r_errs))
        case("readme_root_missing", any(e.startswith("inputs/README.md") for e in r_errs), str(r_errs))
        errs_a, _, _ = check_module(root / "scripts/pkg/a.py", root, counts)
        case("long_public_def_required", any("long_fn" in e for e in errs_a), str(errs_a))
        case("short_public_def_not_required", not any("short_fn" in e for e in errs_a), str(errs_a))
        errs_b, _, _ = check_module(root / "scripts/pkg/b.py", root, counts)
        case("authorship_only_module_docstring", any("authorship-only" in e for e in errs_b), str(errs_b))
        errs_u, _, _ = check_module(root / "scripts/lib/util.py", root, counts)
        names = " ".join(errs_u)
        case("library_all_public_required", "pub" in names and "K" in names and "K.m" in names, names)
        case("library_private_and_nested_ignored", "_priv" not in names and "inner" not in names and "_h" not in names, names)
        case("imported_by_two_is_library", is_library(root / "scripts/nolib/shared.py", root, counts))
        case("imported_by_none_is_not_library", not is_library(root / "scripts/nolib/x.py", root, counts))
        # strip_docstrings / compare logic
        t1 = ast.dump(strip_docstrings(ast.parse('"""a"""\ndef f():\n    """d"""\n    return 1\n')))
        t2 = ast.dump(strip_docstrings(ast.parse('def f():\n    return 1\n')))
        t3 = ast.dump(strip_docstrings(ast.parse('def f():\n    return 2\n')))
        case("strip_docstrings_equal_when_only_docs_differ", t1 == t2)
        case("strip_docstrings_detects_logic_change", t1 != t3)
        # compare_ast on a git repo with one new file and one docstring-only edit
        git_ok = True
        try:
            subprocess.run(["git", "-C", tmp, "init", "-q"], check=True, capture_output=True)
            subprocess.run(["git", "-C", tmp, "-c", "user.email=t@t", "-c", "user.name=t", "add", "-A"], check=True, capture_output=True)
            subprocess.run(["git", "-C", tmp, "-c", "user.email=t@t", "-c", "user.name=t", "commit", "-qm", "base"], check=True, capture_output=True)
        except (subprocess.CalledProcessError, FileNotFoundError):
            git_ok = False
        if git_ok:
            _write(root / "scripts/nolib/x.py", '"""X changed doc only."""\nimport shared\ndef f():\n    """New."""\n    return 1\n')
            _write(root / "scripts/nolib/new.py", '"""New file."""\n')
            _write(root / "scripts/nolib/shared.py", '"""Shared."""\ndef g():\n    return 2\n')
            subprocess.run(["git", "-C", tmp, "add", "-N", "scripts/nolib/new.py"], check=True, capture_output=True)
            f = compare_ast(root, "HEAD", CODE_ROOTS)
            case("compare_ast_new_file (skipped, no crash)", not any("new.py" in e for e in f), str(f))
            case("compare_ast_docstring_only_passes", not any("x.py" in e for e in f), str(f))
            case("compare_ast_logic_change_detected", any("shared.py" in e for e in f), str(f))
        else:
            print("[skip] compare_ast cases (git not available)")
    CODE_ROOTS, README_ROOTS, COVERED_BY_PARENT, LIBRARY_DIRS, EXEMPT_FILES = saved
    print("selftest:", "OK" if failures == 0 else f"{failures} FAILED")
    return 1 if failures else 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--summary", action="store_true", help="print coverage per folder, exit 0")
    ap.add_argument("--paths", nargs="*", default=[], help="limit to these repo-relative paths")
    ap.add_argument("--compare-ast", metavar="REF", help="fail if any changed .py differs from REF beyond docstrings")
    ap.add_argument("--selftest", action="store_true", help="run the built-in fixture tests")
    args = ap.parse_args(argv)
    if args.selftest:
        return selftest()
    root = repo_root()
    if args.compare_ast:
        findings = compare_ast(root, args.compare_ast, CODE_ROOTS)
        for f in findings:
            print(f)
        print("compare-ast:", "OK (docstrings only)" if not findings else f"{len(findings)} file(s) with logic changes")
        return 1 if findings else 0
    return run(root, args.paths, args.summary)


if __name__ == "__main__":
    sys.exit(main())

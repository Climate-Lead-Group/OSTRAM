# run.py
# -*- coding: utf-8 -*-
"""
DVC runner for Windows with Conda environment management and temporary dvc.yaml patching.

Author: Climate Lead Group, Andrey Salazar-Vargas

Features:
- Backup dvc.yaml, temporary replacement of 'DATEPLACEHOLDER' -> YYYY-MM-DD (any occurrence).
- If the Conda environment already exists, it is NOT recreated.
- If the environment exists, checks dependencies and installs missing ones:
    * conda-forge: pandas, numpy, openpyxl, pyyaml, xlsxwriter
    * pip: dvc, otoole
  (installs 'pip' in the environment if necessary).
- Initializes the DVC repository if it doesn't exist (.dvc/).
- Runs 'dvc pull' only if a remote is configured.
- Runs 'dvc repro'.
- Restores dvc.yaml from backup and DELETES the .bak file (always).
"""

import argparse
import datetime as dt
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
import json

# ---------- Default config ----------
ENV_NAME_DEFAULT = "OSTRAM-env"
ENV_FILE_DEFAULT = "environment.yaml"
DVC_FILE_DEFAULT = "dvc.yaml"

# Dependencies to check/install
CONDA_DEPS = {
    # python_module: conda_package
    "pandas": "pandas",
    "numpy": "numpy",
    "openpyxl": "openpyxl",
    "yaml": "pyyaml",          # PyYAML is imported as 'yaml'
    "xlsxwriter": "xlsxwriter"
}
PIP_DEPS = {
    # python_module: pip_package
    "dvc": "dvc",
    "otoole": "otoole>=1.1.1",
}

# ---------- Shell utilities ----------
def run(cmd: str) -> None:
    # Set PYTHONHASHSEED for deterministic hash-based operations
    env = os.environ.copy()
    env['PYTHONHASHSEED'] = '0'
    subprocess.check_call(cmd, shell=True, env=env)

def check_tool_available(tool: str) -> None:
    try:
        subprocess.check_call(f"{tool} --version", shell=True,
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as exc:
        raise RuntimeError(
            f"Requirement '{tool}' not found in PATH. "
            f"Open an Anaconda/Miniconda Prompt or install the tool. Original error: {exc}"
        )

# ---------- Conda environment management ----------
def env_exists(name: str) -> bool:
    """
    Returns True if a conda environment exists whose final directory matches 'name'.
    E.g.: .../envs/OSTRAM-env  -> name == 'OSTRAM-env'
    Uses 'conda env list --json' with text parsing fallback.
    """
    target = name.lower()

    # 1) Primary path: JSON
    try:
        out = subprocess.check_output(
            ["conda", "env", "list", "--json"],
            text=True,
            stderr=subprocess.STDOUT
        )
        data = json.loads(out)
        envs = data.get("envs", []) or []
        return any(Path(p).name.lower() == target for p in envs)
    # If conda is too old or something fails, fall back to text parsing
    except Exception:
        pass

    # 2) Fallback: text parsing of 'conda env list'
    try:
        txt = subprocess.check_output(
            ["conda", "env", "list"],
            text=True,
            stderr=subprocess.STDOUT
        )
        for line in txt.splitlines():
            line = line.strip()
            if not line or line.startswith(("#", "conda environments:")):
                continue
            # typical lines:
            # base                  *  C:\Users\...\anaconda3
            # OSTRAM-env              C:\Users\...\envs\OSTRAM-env
            parts = line.split()
            if not parts:
                continue
            # If the second column is '*', the first is the name
            cand = parts[0].lower()
            if cand == target:
                return True
        return False
    except Exception:
        return False


def guess_env_name_from_yaml(env_file: str) -> str | None:
    p = Path(env_file)
    if not p.exists():
        return None
    try:
        # Simple parsing: look for 'name: ...' line
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.lower().startswith("name:"):
                val = line.split(":", 1)[1].strip().strip("'\"")
                return val or None
    except Exception:
        pass
    return None

def create_env_if_missing(env_name: str, env_file: str) -> None:
    if env_exists(env_name):
        print(f"Conda environment '{env_name}' already exists. Skipping recreation.")
        return
    print(f"Creating Conda environment '{env_name}' from {env_file}...")
    run(f"conda env create -n {env_name} -f {env_file} -y")

def ensure_pip_available(env_name: str) -> None:
    try:
        run(f"conda run -n {env_name} python -m pip --version")
    except subprocess.CalledProcessError:
        print("pip not found in the environment. Installing 'pip' in the environment...")
        run(f"conda install -n {env_name} pip -y")

def module_present(env_name: str, module: str) -> bool:
    code = (
        "import importlib.util,sys;"
        f"sys.exit(0) if importlib.util.find_spec('{module}') else sys.exit(1)"
    )
    try:
        run(f'conda run -n {env_name} python -c "{code}"')
        return True
    except subprocess.CalledProcessError:
        return False

def ensure_deps(env_name: str) -> None:
    """
    Checks modules in the environment and installs missing ones.
    - Conda (conda-forge) for the data stack.
    - Pip for dvc/otoole.
    """
    # Ensure pip is available in the environment if we need it
    need_pip = any(not module_present(env_name, m) for m in list(PIP_DEPS.keys()))
    if need_pip:
        ensure_pip_available(env_name)

    # Conda dependencies
    missing_conda = [pkg for mod, pkg in CONDA_DEPS.items() if not module_present(env_name, mod)]
    if missing_conda:
        pkgs = " ".join(missing_conda)
        print(f"Installing missing conda dependencies: {missing_conda}")
        run(f"conda install -n {env_name} -c conda-forge -y {pkgs}")

    # Pip dependencies
    missing_pip = [pkg for mod, pkg in PIP_DEPS.items() if not module_present(env_name, mod)]
    if missing_pip:
        for spec in missing_pip:
            print(f"Installing missing pip dependency: {spec}")
            run(f"conda run -n {env_name} python -m pip install -U {spec}")

# ---------- DVC ----------
def is_dvc_repo() -> bool:
    return (Path(".dvc").is_dir())

def is_git_repo() -> bool:
    """Checks if the current directory is inside a Git repository."""
    return (Path(".git").is_dir())

def ensure_dvc_repo(env_name: str) -> None:
    if is_dvc_repo():
        print("DVC repository detected (.dvc/ found).")
        return

    # Check if we're in a Git repo to decide how to initialize DVC
    if is_git_repo():
        print("DVC repo not found. Running `dvc init`...")
        run(f"conda run -n {env_name} dvc init")
    else:
        print("Git repo not found. Running `dvc init --no-scm`...")
        run(f"conda run -n {env_name} dvc init --no-scm")

    if not is_dvc_repo():
        raise RuntimeError("Failed to initialize DVC (.dvc was not created).")

def has_dvc_remote(env_name: str) -> bool:
    try:
        out = subprocess.check_output(f"conda run -n {env_name} dvc remote list",
                                      shell=True, stderr=subprocess.STDOUT)
        return bool(out.decode("utf-8", errors="ignore").strip())
    except subprocess.CalledProcessError:
        return False

def dvc_command(env_name: str, args: str) -> None:
    run(f"conda run -n {env_name} dvc {args}")

# ---------- Backup / patching of dvc.yaml ----------
def backup_file(src: Path) -> Path:
    if not src.exists():
        raise FileNotFoundError(f"{src} not found for backup.")
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = src.with_suffix(src.suffix + f".bak.{ts}")
    shutil.copy2(src, bak)
    print(f"Backup created: {bak.name}")
    return bak

def restore_and_delete_backup(backup_path: Path, target: Path) -> None:
    if backup_path and backup_path.exists():
        shutil.copy2(backup_path, target)
        print(f"Restored {target.name} from backup: {backup_path.name}")
        try:
            backup_path.unlink()  # delete the .bak
            print(f"Backup deleted: {backup_path.name}")
        except Exception as e:
            print(f"Warning: could not delete backup ({e})")
    else:
        print("Backup not found; nothing to restore/delete.")

def patch_date_placeholder(dvc_path: Path, date_stamp: str) -> int:
    """
    Replaces ALL literal occurrences of 'DATEPLACEHOLDER' with the date (YYYY-MM-DD).
    Does not use regex, so it also covers '..._DATEPLACEHOLDER.csv' (with underscore).
    """
    text = dvc_path.read_text(encoding="utf-8")
    count = text.count("DATEPLACEHOLDER")
    if count:
        dvc_path.write_text(text.replace("DATEPLACEHOLDER", date_stamp), encoding="utf-8", newline="\n")
    return count

# ---------- Main ----------
def main():
    parser = argparse.ArgumentParser(description="DVC runner with Conda environment and temporary dvc.yaml patching")
    parser.add_argument("--env-name", default=None, help="Conda environment name (if not provided, tries to read it from YAML).")
    parser.add_argument("--env-file", default=ENV_FILE_DEFAULT, help="Path to environment.yaml.")
    parser.add_argument("--dvc-file", default=DVC_FILE_DEFAULT, help="Path to dvc.yaml to patch.")
    parser.add_argument("--date", default=None, help="Date YYYY-MM-DD (defaults to today).")
    args = parser.parse_args()

    # Determine env_name
    env_name = args.env_name or guess_env_name_from_yaml(args.env_file) or ENV_NAME_DEFAULT
    env_file = args.env_file
    dvc_file = Path(args.dvc_file).resolve()

    # Date
    if args.date:
        try:
            dt.date.fromisoformat(args.date)
        except ValueError:
            raise SystemExit("Invalid --date format. Use YYYY-MM-DD (e.g. 2025-08-21).")
        date_stamp = args.date
    else:
        date_stamp = dt.date.today().isoformat()

    print(f"Using environment: {env_name}")
    print(f"Using date: {date_stamp}")
    print(f"dvc.yaml: {dvc_file}")

    # Base requirements
    check_tool_available("conda")

    backup_path = None
    try:
        # 1) Backup + patch 'DATEPLACEHOLDER' before everything else
        backup_path = backup_file(dvc_file)
        replaced = patch_date_placeholder(dvc_file, date_stamp)
        if replaced:
            print(f"Patch applied: {replaced} occurrence(s) of 'DATEPLACEHOLDER' replaced with '{date_stamp}'.")
        else:
            print("No occurrences of 'DATEPLACEHOLDER' found in dvc.yaml.")

        # 2) Environment: create if missing; if it already exists, do NOT recreate.
        create_env_if_missing(env_name, env_file)

        # 3) Check/install dependencies within the environment
        ensure_deps(env_name)

        # 4) Ensure DVC repo
        ensure_dvc_repo(env_name)

        # 5) Pull only if a remote is configured
        if has_dvc_remote(env_name):
            print("📥 dvc pull…")
            dvc_command(env_name, "pull")
        else:
            print("No DVC remote configured. Skipping `dvc pull`.")

        # 6) Reproduce pipeline
        print("🔄 dvc repro…")
        start_time = dt.datetime.now()
        dvc_command(env_name, "repro")
        end_time = dt.datetime.now()

        # Calculate and display duration
        duration = end_time - start_time
        total_seconds = int(duration.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        duration_str = []
        if hours > 0:
            duration_str.append(f"{hours}h")
        if minutes > 0 or hours > 0:
            duration_str.append(f"{minutes}m")
        duration_str.append(f"{seconds}s")

        print(f"Pipeline completed in {' '.join(duration_str)}!")

    finally:
        # 7) Restore dvc.yaml and delete backup
        if backup_path:
            restore_and_delete_backup(backup_path, dvc_file)

if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as e:
        print(f"\nCommand failed (exit {e.returncode}): {e.cmd}", file=sys.stderr)
        sys.exit(e.returncode)
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        sys.exit(1)

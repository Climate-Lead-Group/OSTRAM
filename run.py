#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Top-level runner for OSTRAM on Windows with Conda environment management.

Author: Climate Lead Group, Andrey Salazar-Vargas

Features:
- Reuses the Conda environment if it already exists.
- Installs missing dependencies into the existing environment.
- Initializes the DVC repository if it does not exist.
- Runs `dvc pull` only when a remote is configured.
- Executes A3 (pre-process A1 outputs), B1, and B2 explicitly from this top-level launcher.
"""

import argparse
import datetime as dt
import json
import os
import subprocess
import sys
from pathlib import Path

# ---------- Default config ----------
ENV_NAME_DEFAULT = "OSTRAM-env"
ENV_FILE_DEFAULT = "environment.yaml"
DVC_FILE_DEFAULT = "dvc.yaml"
T1_DIR = Path("t1_confection")
A3_SCRIPT_DEFAULT = T1_DIR / "A3_process.py"
B1_SCRIPT_DEFAULT = T1_DIR / "B1_Run_Compiler.py"
B2_SCRIPT_DEFAULT = T1_DIR / "B2_Executing_OG_Model.py"

# Dependencies to check/install
CONDA_DEPS = {
    "pandas": "pandas",
    "numpy": "numpy",
    "openpyxl": "openpyxl",
    "yaml": "pyyaml",
    "ruamel.yaml": "ruamel.yaml",
    "xlsxwriter": "xlsxwriter",
}
PIP_DEPS = {
    "dvc": "dvc",
    "otoole": "otoole>=1.1.1",
}


# ---------- Shell utilities ----------
def run(cmd: str) -> None:
    env = os.environ.copy()
    env["PYTHONHASHSEED"] = "0"
    subprocess.check_call(cmd, shell=True, env=env)


def check_tool_available(tool: str) -> None:
    try:
        subprocess.check_call(
            f"{tool} --version",
            shell=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        raise RuntimeError(
            f"Requirement '{tool}' not found in PATH. "
            f"Open an Anaconda/Miniconda Prompt or install the tool. Original error: {exc}"
        )


# ---------- Conda environment management ----------
def env_exists(name: str) -> bool:
    target = name.lower()

    try:
        out = subprocess.check_output(
            ["conda", "env", "list", "--json"],
            text=True,
            stderr=subprocess.STDOUT,
        )
        data = json.loads(out)
        envs = data.get("envs", []) or []
        return any(Path(p).name.lower() == target for p in envs)
    except Exception:
        pass

    try:
        txt = subprocess.check_output(
            ["conda", "env", "list"],
            text=True,
            stderr=subprocess.STDOUT,
        )
        for line in txt.splitlines():
            line = line.strip()
            if not line or line.startswith(("#", "conda environments:")):
                continue
            parts = line.split()
            if parts and parts[0].lower() == target:
                return True
        return False
    except Exception:
        return False


def guess_env_name_from_yaml(env_file: str) -> str | None:
    p = Path(env_file)
    if not p.exists():
        return None
    try:
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
    need_pip = any(not module_present(env_name, m) for m in PIP_DEPS.keys())
    if need_pip:
        ensure_pip_available(env_name)

    missing_conda = [pkg for mod, pkg in CONDA_DEPS.items() if not module_present(env_name, mod)]
    if missing_conda:
        pkgs = " ".join(missing_conda)
        print(f"Installing missing conda dependencies: {missing_conda}")
        run(f"conda install -n {env_name} -c conda-forge -y {pkgs}")

    missing_pip = [pkg for mod, pkg in PIP_DEPS.items() if not module_present(env_name, mod)]
    if missing_pip:
        for spec in missing_pip:
            print(f"Installing missing pip dependency: {spec}")
            run(f"conda run -n {env_name} python -m pip install -U {spec}")


# ---------- DVC ----------
def is_dvc_repo() -> bool:
    return Path(".dvc").is_dir()


def is_git_repo() -> bool:
    return Path(".git").is_dir()


def ensure_dvc_repo(env_name: str) -> None:
    if is_dvc_repo():
        print("DVC repository detected (.dvc/ found).")
        return

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
        out = subprocess.check_output(
            f"conda run -n {env_name} dvc remote list",
            shell=True,
            stderr=subprocess.STDOUT,
        )
        return bool(out.decode("utf-8", errors="ignore").strip())
    except subprocess.CalledProcessError:
        return False


def dvc_command(env_name: str, args: str) -> None:
    run(f"conda run -n {env_name} dvc {args}")


def run_pipeline_script(env_name: str, script_path: Path) -> None:
    if not script_path.is_file():
        raise FileNotFoundError(f"Pipeline script not found: {script_path}")

    print(f"Running {script_path.relative_to(Path.cwd())}...")
    run(f'conda run -n {env_name} python -u "{script_path}"')


def format_duration(start_time: dt.datetime, end_time: dt.datetime) -> str:
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
    return " ".join(duration_str)


# ---------- Main ----------
def main() -> None:
    parser = argparse.ArgumentParser(description="Top-level runner for OSTRAM A3/B1/B2 execution")
    parser.add_argument(
        "--env-name",
        default=None,
        help="Conda environment name (if not provided, tries to read it from YAML).",
    )
    parser.add_argument("--env-file", default=ENV_FILE_DEFAULT, help="Path to environment.yaml.")
    parser.add_argument(
        "--dvc-file",
        default=DVC_FILE_DEFAULT,
        help="Path to dvc.yaml used for optional DVC pull checks.",
    )
    parser.add_argument("--skip-pull", action="store_true", help="Skip `dvc pull` even if a DVC remote is configured.")
    parser.add_argument("--skip-a3", action="store_true", help="Skip `t1_confection/A3_process.py` (pre-process A1 outputs).")
    parser.add_argument("--skip-b1", action="store_true", help="Skip `t1_confection/B1_Run_Compiler.py`.")
    parser.add_argument("--skip-b2", action="store_true", help="Skip `t1_confection/B2_Executing_OG_Model.py`.")
    args = parser.parse_args()

    env_name = args.env_name or guess_env_name_from_yaml(args.env_file) or ENV_NAME_DEFAULT
    env_file = args.env_file
    dvc_file = Path(args.dvc_file).resolve()

    print(f"Using environment: {env_name}")
    print(f"DVC config: {dvc_file}")

    check_tool_available("conda")

    create_env_if_missing(env_name, env_file)
    ensure_deps(env_name)
    ensure_dvc_repo(env_name)

    if args.skip_pull:
        print("Skipping `dvc pull` by request.")
    elif has_dvc_remote(env_name):
        print("Pulling DVC data...")
        dvc_command(env_name, "pull")
    else:
        print("No DVC remote configured. Skipping `dvc pull`.")

    start_time = dt.datetime.now()

    if args.skip_a3:
        print("Skipping A3 pre-process stage by request.")
    else:
        run_pipeline_script(env_name, A3_SCRIPT_DEFAULT.resolve())

    if args.skip_b1:
        print("Skipping B1 compiler stage by request.")
    else:
        run_pipeline_script(env_name, B1_SCRIPT_DEFAULT.resolve())

    if args.skip_b2:
        print("Skipping B2 execution stage by request.")
    else:
        run_pipeline_script(env_name, B2_SCRIPT_DEFAULT.resolve())

    end_time = dt.datetime.now()
    print(f"Pipeline completed in {format_duration(start_time, end_time)}.")


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as e:
        print(f"\nCommand failed (exit {e.returncode}): {e.cmd}", file=sys.stderr)
        sys.exit(e.returncode)
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        sys.exit(1)

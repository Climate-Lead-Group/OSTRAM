# -*- coding: utf-8 -*-
"""
Canonical pipeline orchestration for OSTRAM with Conda environment management.

Author: Climate Lead Group, Andrey Salazar-Vargas

Features:
- Reuses the Conda environment if it already exists.
- Installs missing dependencies into the existing environment.
- Initializes the DVC repository if it does not exist unless `--skip-pull` is
  selected.
- Runs `dvc pull` only when a remote is configured.
- If any required canonical root lacks a post-A2 snapshot, runs A1 + A2 as a
  combo first. Then materializes one exact registry-selected scenario set and
  routes that same ordered set through B1 and B2.
"""

import argparse
import datetime as dt
import json
import os
import shlex
import subprocess
import sys
from pathlib import Path
from typing import Sequence

from ostram.paths import ProjectPaths, resolve_paths

from ostram.pipeline.scenarios.registry import (
    ensure_root_output_directories,
    load_registry,
    root_snapshots_exist,
)

# ---------- Default config ----------
ENV_NAME_DEFAULT = "OSTRAM-env"
ENV_FILE_DEFAULT = "environment.yaml"
DVC_FILE_DEFAULT = "dvc.yaml"
_DEFAULT_PATHS = resolve_paths()
PIPELINE_DIR = _DEFAULT_PATHS.package_root / "pipeline"
A1_SCRIPT_DEFAULT = PIPELINE_DIR / "preparation" / "base_inputs.py"
A2_SCRIPT_DEFAULT = PIPELINE_DIR / "preparation" / "transmission.py"
A3_SCRIPT_DEFAULT = PIPELINE_DIR / "scenarios" / "materializer.py"
B1_SCRIPT_DEFAULT = PIPELINE_DIR / "compilation" / "runner.py"
B2_SCRIPT_DEFAULT = PIPELINE_DIR / "execution" / "runner.py"
A1_OUTPUTS_DIR = _DEFAULT_PATHS.a1_outputs
SNAPSHOT_PREFIX = "_post_a2_snapshot_"

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
def run(cmd: Sequence[str | Path], *, cwd: Path | None = None) -> None:
    env = os.environ.copy()
    env["PYTHONHASHSEED"] = "0"
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    tokens = shlex.split(cmd) if isinstance(cmd, str) else list(cmd)
    subprocess.check_call(
        [str(token) for token in tokens],
        cwd=str(cwd.resolve()) if cwd is not None else None,
        env=env,
    )


def check_tool_available(tool: str) -> None:
    try:
        subprocess.check_call(
            [tool, "--version"],
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


def guess_env_name_from_yaml(env_file: str | Path) -> str | None:
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


def create_env_if_missing(env_name: str, env_file: str | Path) -> None:
    if env_exists(env_name):
        print(f"Conda environment '{env_name}' already exists. Skipping recreation.")
        return
    print(f"Creating Conda environment '{env_name}' from {env_file}...")
    run(["conda", "env", "create", "-n", env_name, "-f", Path(env_file), "-y"])


def ensure_pip_available(env_name: str) -> None:
    try:
        run(["conda", "run", "-n", env_name, "python", "-m", "pip", "--version"])
    except subprocess.CalledProcessError:
        print("pip not found in the environment. Installing 'pip' in the environment...")
        run(["conda", "install", "-n", env_name, "pip", "-y"])


def module_present(env_name: str, module: str) -> bool:
    code = (
        "import importlib.util,sys;"
        f"sys.exit(0) if importlib.util.find_spec('{module}') else sys.exit(1)"
    )
    try:
        run(["conda", "run", "-n", env_name, "python", "-c", code])
        return True
    except subprocess.CalledProcessError:
        return False


def ensure_deps(env_name: str) -> None:
    need_pip = any(not module_present(env_name, m) for m in PIP_DEPS.keys())
    if need_pip:
        ensure_pip_available(env_name)

    missing_conda = [pkg for mod, pkg in CONDA_DEPS.items() if not module_present(env_name, mod)]
    if missing_conda:
        print(f"Installing missing conda dependencies: {missing_conda}")
        run(["conda", "install", "-n", env_name, "-c", "conda-forge", "-y", *missing_conda])

    missing_pip = [pkg for mod, pkg in PIP_DEPS.items() if not module_present(env_name, mod)]
    if missing_pip:
        for spec in missing_pip:
            print(f"Installing missing pip dependency: {spec}")
            run(["conda", "run", "-n", env_name, "python", "-m", "pip", "install", "-U", spec])


# ---------- DVC ----------
def is_dvc_repo(project_root: Path | None = None) -> bool:
    root = resolve_paths().project_root if project_root is None else project_root
    return (root / ".dvc").is_dir()


def is_git_repo(project_root: Path | None = None) -> bool:
    root = resolve_paths().project_root if project_root is None else project_root
    return (root / ".git").exists()


def ensure_dvc_repo(env_name: str, project_root: Path | None = None) -> None:
    root = resolve_paths().project_root if project_root is None else project_root.resolve()
    if is_dvc_repo(root):
        print("DVC repository detected (.dvc/ found).")
        return

    if is_git_repo(root):
        print("DVC repo not found. Running `dvc init`...")
        run(["conda", "run", "-n", env_name, "dvc", "init"], cwd=root)
    else:
        print("Git repo not found. Running `dvc init --no-scm`...")
        run(["conda", "run", "-n", env_name, "dvc", "init", "--no-scm"], cwd=root)

    if not is_dvc_repo(root):
        raise RuntimeError("Failed to initialize DVC (.dvc was not created).")


def has_dvc_remote(env_name: str, project_root: Path | None = None) -> bool:
    root = resolve_paths().project_root if project_root is None else project_root.resolve()
    try:
        out = subprocess.check_output(
            ["conda", "run", "-n", env_name, "dvc", "remote", "list"],
            cwd=str(root),
            stderr=subprocess.STDOUT,
        )
        return bool(out.decode("utf-8", errors="ignore").strip())
    except subprocess.CalledProcessError:
        return False


def dvc_command(
    env_name: str,
    args: Sequence[str] | str,
    project_root: Path | None = None,
) -> None:
    root = resolve_paths().project_root if project_root is None else project_root.resolve()
    tokens = shlex.split(args) if isinstance(args, str) else list(args)
    run(["conda", "run", "-n", env_name, "dvc", *tokens], cwd=root)


def _module_for_script(script_path: Path, paths: ProjectPaths) -> str:
    relative = script_path.resolve().relative_to(paths.project_root)
    return ".".join(relative.with_suffix("").parts)


def run_pipeline_script(
    env_name: str,
    script_path: Path,
    extra_args: Sequence[str] | str = (),
    *,
    paths: ProjectPaths | None = None,
) -> None:
    del env_name  # the active interpreter is the supported package environment
    active_paths = resolve_paths() if paths is None else paths
    if not script_path.is_file():
        raise FileNotFoundError(f"Pipeline script not found: {script_path}")
    arguments = shlex.split(extra_args) if isinstance(extra_args, str) else list(extra_args)
    module = _module_for_script(script_path, active_paths)
    relative = script_path.resolve().relative_to(active_paths.project_root)
    stage = relative.parts[2] if len(relative.parts) > 2 else "pipeline"
    stage_cwd = active_paths.stage_workspace(stage, create=True)
    print(f"Running {module} from {stage_cwd}...")
    run(
        [sys.executable, "-B", "-m", module, *arguments],
        cwd=stage_cwd,
    )


def enumerate_active_scenarios(env_name: str | None = None) -> list[str]:
    """Return BAU plus the frozen decision set from the runtime registry."""

    del env_name  # retained for compatibility with existing callers
    return list(load_registry().select(None))


def run_a3_for_scenario(env_name: str, script_path: Path, scenario: str) -> None:
    """Compatibility helper: materialize one canonical scenario."""
    if not script_path.is_file():
        raise FileNotFoundError(f"scenario materializer not found: {script_path}")
    print(f"Running A3 for scenario '{scenario}'...")
    run_pipeline_script(
        env_name,
        script_path,
        ["--scenarios", scenario],
    )


def run_a3_for_scenarios(
    env_name: str,
    script_path: Path,
    scenarios: list[str],
    a_result_seed: str | None = None,
) -> None:
    """Invoke the canonical materializer once for one exact ordered set."""

    if not script_path.is_file():
        raise FileNotFoundError(f"scenario materializer not found: {script_path}")
    if not scenarios:
        raise ValueError("scenario selection is empty")
    selected = ",".join(scenarios)
    arguments = ["--scenarios", selected]
    if a_result_seed:
        arguments.extend(["--a-result-seed", str(Path(a_result_seed).resolve())])
    print(f"Materializing scenarios in canonical order: {scenarios}")
    run_pipeline_script(env_name, script_path, arguments)


def post_a2_snapshot_exists(a1_outputs_dir: Path) -> bool:
    """True if at least one `_post_a2_snapshot_*` folder exists in A1_Outputs/."""
    if not a1_outputs_dir.is_dir():
        return False
    return any(
        p.is_dir() and p.name.startswith(SNAPSHOT_PREFIX)
        for p in a1_outputs_dir.iterdir()
    )


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
def parse_args(argv=None):
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
    parser.add_argument(
        "--skip-pull",
        action="store_true",
        help="Skip DVC repository setup and `dvc pull`.",
    )
    parser.add_argument("--skip-a3", action="store_true", help="Skip A3 scenario materialization.")
    parser.add_argument("--skip-b1", action="store_true", help="Skip B1 input compilation.")
    parser.add_argument("--skip-b2", action="store_true", help="Skip B2 execution preparation.")
    parser.add_argument(
        "--scenarios",
        default=None,
        help="Comma-separated list of scenarios to run (e.g. "
             "'B_Optimised_VRE,C_Target_VRE'). When omitted, runs all active "
             "scenarios. Filter is propagated to A3, B1 and B2.",
    )
    parser.add_argument(
        "--a-result-seed",
        default=None,
        help=(
            "Read-only A_Calibrated_BAU result CSV or directory used only "
            "for the declared C_Target_VRE dependency."
        ),
    )
    parser.add_argument(
        "--compile-only",
        action="store_true",
        help=(
            "Run A3, B1, and B2 input compilation but stop before every "
            "matrix/solver/output boundary."
        ),
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    paths = resolve_paths()
    env_file = paths.resolve_project_file(args.env_file)
    dvc_file = paths.resolve_project_file(args.dvc_file)
    env_name = args.env_name or guess_env_name_from_yaml(env_file) or ENV_NAME_DEFAULT

    print(f"Using environment: {env_name}")
    print(f"DVC config: {dvc_file}")

    check_tool_available("conda")

    create_env_if_missing(env_name, env_file)
    ensure_deps(env_name)

    if args.skip_pull:
        print("Skipping DVC repository setup and `dvc pull` by request.")
    else:
        ensure_dvc_repo(env_name)
        if has_dvc_remote(env_name):
            print("Pulling DVC data...")
            dvc_command(env_name, "pull")
        else:
            print("No DVC remote configured. Skipping `dvc pull`.")

    start_time = dt.datetime.now()

    registry = load_registry()
    try:
        scenarios = list(registry.select(args.scenarios))
    except ValueError as error:
        raise RuntimeError(str(error)) from error
    if not scenarios:
        raise RuntimeError("--scenarios selected no canonical scenarios")
    required_roots = registry.required_roots(scenarios)

    # A1 + A2 are a combo: every requested root must have its own post-A2
    # snapshot. A1 creates only the four root output directories; A2 snapshots
    # only those roots.
    a1_outputs_dir = paths.a1_outputs
    if root_snapshots_exist(a1_outputs_dir, required_roots):
        print(
            f"All required post-A2 root snapshots exist in {A1_OUTPUTS_DIR}/: "
            f"{list(required_roots)}. Skipping A1 + A2."
        )
    else:
        print(
            f"One or more required post-A2 root snapshots are absent in "
            f"{A1_OUTPUTS_DIR}/. Running A1 + A2 for the four roots."
        )
        ensure_root_output_directories(a1_outputs_dir, registry)
        run_pipeline_script(env_name, A1_SCRIPT_DEFAULT)
        run_pipeline_script(env_name, A2_SCRIPT_DEFAULT)

    if args.skip_a3:
        print("Skipping A3 pre-process stage by request.")
    else:
        run_a3_for_scenarios(
            env_name,
            A3_SCRIPT_DEFAULT,
            scenarios,
            args.a_result_seed,
        )

    scenarios_arg = ["--scenarios", ",".join(scenarios)]

    if args.skip_b1:
        print("Skipping B1 compiler stage by request.")
    else:
        run_pipeline_script(env_name, B1_SCRIPT_DEFAULT, scenarios_arg)

    if args.skip_b2:
        print("Skipping B2 execution stage by request.")
    else:
        b2_args = [*scenarios_arg, *( ["--compile-only"] if args.compile_only else [])]
        run_pipeline_script(env_name, B2_SCRIPT_DEFAULT, b2_args)

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

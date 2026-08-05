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
from ostram.profiles import PROFILE_MANIFEST_ENV, active_profile_id
from ostram.terminal import RunReporter, activate_reporter, active_reporter

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
    # Pipeline stages run from isolated stage workspaces, so ``sys.path[0]``
    # is not the project root.  Put the activated bundle first to prevent an
    # unrelated editable/install copy of ``ostram`` from servicing children.
    project_root = str(resolve_paths().project_root)
    inherited = [
        item
        for item in env.get("PYTHONPATH", "").split(os.pathsep)
        if item and os.path.normcase(os.path.abspath(item))
        != os.path.normcase(os.path.abspath(project_root))
    ]
    env["PYTHONPATH"] = os.pathsep.join([project_root, *inherited])
    tokens = shlex.split(cmd) if isinstance(cmd, str) else list(cmd)
    reporter = active_reporter()
    if reporter is not None:
        reporter.run_child(
            [str(token) for token in tokens],
            cwd=cwd,
            env=env,
        )
        return
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
    parser = argparse.ArgumentParser(
        description=(
            "Run OSTRAM through base-model preparation (A1), transmission "
            "network preparation (A2), scenario building (A3), input "
            "compilation (B1), and model execution/result collection (B2)."
        )
    )
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
    parser.add_argument(
        "--verbose",
        action="store_true",
        help=(
            "Stream complete child-process output and command diagnostics; "
            "the detailed UTF-8 run log is always written."
        ),
    )
    return parser.parse_args(argv)


def _create_run_reporter(
    paths: ProjectPaths,
    scenarios: Sequence[str],
    *,
    verbose: bool,
    compile_only: bool = False,
) -> RunReporter:
    manifest_value = os.environ.get(PROFILE_MANIFEST_ENV)
    return RunReporter(
        project_root=paths.project_root,
        workspace=paths.workspace,
        scenarios=scenarios,
        verbose=verbose,
        profile_id=active_profile_id(),
        manifest=Path(manifest_value) if manifest_value else None,
        compile_only=compile_only,
    )


def main() -> None:
    args = parse_args()
    paths = resolve_paths()
    pipeline_dir = paths.package_root / "pipeline"
    a1_script = pipeline_dir / "preparation" / "base_inputs.py"
    a2_script = pipeline_dir / "preparation" / "transmission.py"
    a3_script = pipeline_dir / "scenarios" / "materializer.py"
    b1_script = pipeline_dir / "compilation" / "runner.py"
    b2_script = pipeline_dir / "execution" / "runner.py"
    env_file = paths.resolve_project_file(args.env_file)
    dvc_file = paths.resolve_project_file(args.dvc_file)
    env_name = args.env_name or guess_env_name_from_yaml(env_file) or ENV_NAME_DEFAULT

    registry = load_registry()
    try:
        scenarios = list(registry.select(args.scenarios))
    except ValueError as error:
        raise RuntimeError(str(error)) from error
    if not scenarios:
        raise RuntimeError("--scenarios selected no canonical scenarios")
    required_roots = registry.required_roots(scenarios)
    reporter = _create_run_reporter(
        paths,
        scenarios,
        verbose=args.verbose,
        compile_only=args.compile_only,
    )
    start_time = dt.datetime.now()

    with activate_reporter(reporter), reporter.capture_output():
        try:
            print(f"Using environment: {env_name}")
            reporter.note(f"Profile: {active_profile_id()}")
            reporter.note(f"Manifest: {os.environ.get(PROFILE_MANIFEST_ENV, 'compatibility default')}")
            reporter.note(f"Workspace: {paths.workspace}")
            print(f"DVC config: {dvc_file}")
            reporter.note(f"Selected scenarios: {', '.join(scenarios)}")

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

            # A1 + A2 are a combo: every requested root must have its own
            # post-A2 snapshot. A1 creates only the four root output
            # directories; A2 snapshots only those roots.
            a1_outputs_dir = paths.a1_outputs
            if root_snapshots_exist(a1_outputs_dir, required_roots):
                reason = "required post-A2 root snapshots already exist"
                reporter.stage_skip("A1", reason)
                reporter.stage_skip("A2", reason)
                print(
                    f"All required post-A2 root snapshots exist in "
                    f"{a1_outputs_dir}/: {list(required_roots)}. "
                    "Skipping base-model and transmission preparation."
                )
            else:
                print(
                    f"One or more required post-A2 root snapshots are absent "
                    f"in {a1_outputs_dir}/. Preparing the four roots."
                )
                reporter.stage_start("A1")
                ensure_root_output_directories(a1_outputs_dir, registry)
                run_pipeline_script(env_name, a1_script)
                reporter.stage_complete("A1")

                reporter.stage_start("A2")
                run_pipeline_script(env_name, a2_script)
                reporter.stage_complete("A2")

            if args.skip_a3:
                reporter.stage_skip("A3", "skipped by --skip-a3")
                print("Skipping scenario building (A3) by request.")
            else:
                current = scenarios[0] if len(scenarios) == 1 else f"{len(scenarios)} selected"
                reporter.stage_start("A3", scenario=current)
                run_a3_for_scenarios(
                    env_name,
                    a3_script,
                    scenarios,
                    args.a_result_seed,
                )
                reporter.stage_complete("A3")

            scenarios_arg = ["--scenarios", ",".join(scenarios)]

            if args.skip_b1:
                reporter.stage_skip("B1", "skipped by --skip-b1")
                print("Skipping model-input compilation (B1) by request.")
            else:
                reporter.stage_start("B1")
                run_pipeline_script(env_name, b1_script, scenarios_arg)
                reporter.stage_complete("B1")

            if args.skip_b2:
                reporter.stage_skip("B2", "skipped by --skip-b2")
                print("Skipping model execution and result collection (B2) by request.")
            else:
                reporter.stage_start("B2")
                b2_args = [
                    *scenarios_arg,
                    *(["--compile-only"] if args.compile_only else []),
                ]
                run_pipeline_script(env_name, b2_script, b2_args)
                reporter.stage_complete(
                    "B2",
                    detail=(
                        "compile-only; matrix and solver skipped"
                        if args.compile_only
                        else None
                    ),
                )

            end_time = dt.datetime.now()
            print(f"Pipeline stages finished in {format_duration(start_time, end_time)}.")
        except KeyboardInterrupt:
            reporter.stage_fail("interrupted by user")
            reporter.finish(
                outcome="INTERRUPTED",
                exit_code=130,
                final_message="OSTRAM run interrupted",
            )
            raise
        except subprocess.CalledProcessError as error:
            reporter.stage_fail(f"child process exited with code {error.returncode}")
            reporter.finish(
                outcome="FAILED",
                exit_code=error.returncode,
                final_message="OSTRAM run failed",
            )
            raise
        except Exception as error:
            reporter.stage_fail(f"{type(error).__name__}: {error}")
            reporter.finish(
                outcome="FAILED",
                exit_code=1,
                final_message="OSTRAM run failed",
            )
            raise
        else:
            compile_only_completed = args.compile_only and not args.skip_b2
            reporter.finish(
                outcome="COMPILE_ONLY_SUCCESS" if compile_only_completed else "SUCCESS",
                exit_code=0,
                final_message=(
                    "OSTRAM compile-only run completed successfully"
                    if compile_only_completed
                    else "OSTRAM run completed successfully"
                ),
            )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(130)
    except subprocess.CalledProcessError as e:
        print(f"\nCommand failed (exit {e.returncode}): {e.cmd}", file=sys.stderr)
        sys.exit(e.returncode)
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        sys.exit(1)

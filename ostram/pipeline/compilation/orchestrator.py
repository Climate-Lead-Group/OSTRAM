"""Behavior-preserving orchestration for the B1 compiler runner."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Callable, Iterator, List, Optional, Sequence

from ostram.paths import resolve_paths


_TIMESTAMP_RE = re.compile(r"_\d{8}")


@dataclass(frozen=True)
class B1Paths:
    """Filesystem paths used by one B1 orchestration run."""

    script_dir: Path
    config_path: Path
    compiler_path: Path
    scenarios_root: Path

    @classmethod
    def defaults(cls) -> "B1Paths":
        import yaml

        project = resolve_paths()
        runtime_root = project.stage_workspace("compilation", create=True)
        runtime_config = runtime_root / "Config_MOMF_T1_A.yaml"
        if not runtime_config.exists():
            shutil.copy2(project.compilation_config, runtime_config)
        data = yaml.safe_load(runtime_config.read_text(encoding="utf-8"))
        data.update(
            {
                "A1_outputs": str(project.a1_outputs),
                "A2_extra_inputs": str(project.generated_extra_inputs),
                "A2_output": str(project.compiled_parameters),
                "A2_output_main_scen": str(project.compiled_parameters),
            }
        )
        runtime_config.write_text(
            yaml.safe_dump(data, sort_keys=False, allow_unicode=True),
            encoding="utf-8",
        )
        script_dir = Path(__file__).resolve().parent
        return cls(
            script_dir=runtime_root,
            config_path=runtime_config,
            compiler_path=script_dir / "compiler.py",
            scenarios_root=project.a1_outputs,
        )


@dataclass(frozen=True)
class ScenarioSelection:
    """Resolved scenario filter while retaining predecessor diagnostics."""

    selected: tuple[str, ...]
    requested: tuple[str, ...]
    unknown: tuple[str, ...]
    filter_active: bool


@dataclass(frozen=True)
class CompilerCommand:
    """A compiler invocation with explicit tokens and working directory."""

    argv: tuple[str, ...]
    cwd: Path


def try_import_yaml_handlers():
    ruamel_yaml = None
    pyyaml = None
    try:
        from ruamel.yaml import YAML  # type: ignore

        ruamel_yaml = YAML
    except Exception:
        ruamel_yaml = None
    if ruamel_yaml is None:
        try:
            import yaml  # type: ignore

            pyyaml = yaml
        except Exception:
            pyyaml = None
    return ruamel_yaml, pyyaml


def list_scenario_suffixes(base_dir: Path) -> List[str]:
    """Return eligible suffixes from sorted ``A1_Outputs_*`` directories."""
    suffixes: List[str] = []
    for item in sorted(base_dir.iterdir()):
        if not (item.is_dir() and item.name.startswith("A1_Outputs_")):
            continue
        suffix = item.name.split("A1_Outputs_", 1)[1]
        if not suffix:
            continue
        suffix_lower = suffix.lower()
        if "backup" in suffix_lower:
            continue
        if "snapshot" in suffix_lower:
            continue
        if "pre_experiment" in suffix_lower:
            continue
        if _TIMESTAMP_RE.search(suffix):
            continue
        suffixes.append(suffix)
    return suffixes


def resolve_scenarios(
    discovered: Sequence[str], raw_filter: Optional[str]
) -> ScenarioSelection:
    """Resolve a CLI filter in its exact requested canonical order."""
    discovered_list = list(discovered)
    if not raw_filter:
        return ScenarioSelection(
            selected=tuple(discovered_list),
            requested=(),
            unknown=(),
            filter_active=False,
        )

    requested = tuple(
        scenario.strip() for scenario in raw_filter.split(",") if scenario.strip()
    )
    requested_unique = tuple(dict.fromkeys(requested))
    discovered_set = set(discovered_list)
    unknown = tuple(
        scenario for scenario in requested if scenario not in discovered_set
    )
    selected = tuple(
        scenario for scenario in requested_unique if scenario in discovered_set
    )
    return ScenarioSelection(
        selected=selected,
        requested=requested,
        unknown=unknown,
        filter_active=True,
    )


def read_yaml_ruamel(yaml_path: Path, YAML_cls):
    """Read YAML using ruamel.yaml to preserve formatting."""
    yaml = YAML_cls()
    yaml.preserve_quotes = True
    with yaml_path.open("r", encoding="utf-8") as file:
        data = yaml.load(file)
    return data, yaml


def write_yaml_ruamel(yaml_path: Path, data, yaml_obj):
    """Write YAML using ruamel.yaml."""
    with yaml_path.open("w", encoding="utf-8") as file:
        yaml_obj.dump(data, file)


def read_yaml_pyyaml(yaml_path: Path, pyyaml):
    """Read YAML using PyYAML (comments will be lost)."""
    with yaml_path.open("r", encoding="utf-8") as file:
        data = pyyaml.safe_load(file)
    return data


def write_yaml_pyyaml(yaml_path: Path, data, pyyaml):
    """Write YAML using PyYAML."""
    with yaml_path.open("w", encoding="utf-8") as file:
        pyyaml.safe_dump(data, file, sort_keys=False, allow_unicode=True)


def regex_update_main_scenario(yaml_text: str, new_value: str) -> str:
    """Preserve the predecessor's last-resort regex update behavior."""
    xtra_match = re.search(r"(^|\n)xtra_scen:\s*\{?[\s\S]*?$", yaml_text)
    if not xtra_match:
        return re.sub(
            r"(Main_Scenario:\s*)['\"]?.*?['\"]?",
            rf"\1'{new_value}'",
            yaml_text,
            count=1,
        )

    def replace_first_after_xtra(text: str) -> str:
        return re.sub(
            r"(Main_Scenario:\s*)['\"]?.*?['\"]?",
            rf"\1'{new_value}'",
            text,
            count=1,
        )

    return replace_first_after_xtra(yaml_text)


def update_main_scenario(yaml_path: Path, new_value: str) -> None:
    """Update ``xtra_scen.Main_Scenario`` with the available YAML handler."""
    ruamel_yaml_cls, pyyaml_mod = try_import_yaml_handlers()

    if ruamel_yaml_cls is not None:
        data, yaml_obj = read_yaml_ruamel(yaml_path, ruamel_yaml_cls)
        if not isinstance(data, dict) or "xtra_scen" not in data:
            raise ValueError("YAML does not contain 'xtra_scen' at the top level.")
        if not isinstance(data["xtra_scen"], dict):
            raise ValueError("'xtra_scen' is not a mapping in the YAML.")
        data["xtra_scen"]["Main_Scenario"] = new_value
        write_yaml_ruamel(yaml_path, data, yaml_obj)
        return

    if pyyaml_mod is not None:
        data = read_yaml_pyyaml(yaml_path, pyyaml_mod)
        if (
            not isinstance(data, dict)
            or "xtra_scen" not in data
            or not isinstance(data["xtra_scen"], dict)
        ):
            raise ValueError("YAML structure invalid or missing 'xtra_scen'.")
        data["xtra_scen"]["Main_Scenario"] = new_value
        write_yaml_pyyaml(yaml_path, data, pyyaml_mod)
        return

    original_text = yaml_path.read_text(encoding="utf-8")
    updated_text = regex_update_main_scenario(original_text, new_value)
    yaml_path.write_text(updated_text, encoding="utf-8")


def build_compiler_command(
    *, interpreter: str, compiler_path: Path, cwd: Path
) -> CompilerCommand:
    """Build the exact tokenized compiler command without executing it."""
    if compiler_path.name not in {"B1_Compiler.py", "compiler.py"}:
        raise ValueError(f"unsupported compiler module target: {compiler_path}")
    return CompilerCommand(
        argv=(interpreter, "-B", "-m", "ostram.pipeline.compilation.compiler"),
        cwd=cwd.resolve(),
    )


def execute_command(
    command: CompilerCommand,
    *,
    command_runner: Optional[Callable[..., object]] = None,
) -> int:
    """Execute a compiler command while inheriting the complete environment."""
    runner = subprocess.run if command_runner is None else command_runner
    result = runner(list(command.argv), cwd=str(command.cwd))
    return result.returncode


def run_compiler(
    script_dir: Path,
    *,
    interpreter: Optional[str] = None,
    command_runner: Optional[Callable[..., object]] = None,
) -> int:
    """Execute ``B1_Compiler.py`` with the selected current interpreter."""
    candidate = script_dir / "compiler.py"
    compiler = (
        Path(__file__).with_name("compiler.py")
        if script_dir.resolve() == resolve_paths().compilation_workspace
        else candidate
    )
    if not compiler.is_file():
        raise FileNotFoundError(f"Missing script: {compiler}")
    command = build_compiler_command(
        interpreter=sys.executable if interpreter is None else interpreter,
        compiler_path=compiler,
        cwd=script_dir,
    )
    return execute_command(command, command_runner=command_runner)


@contextmanager
def preserved_configuration(
    yaml_path: Path,
    *,
    copy_file: Optional[Callable[..., object]] = None,
    move_file: Optional[Callable[..., object]] = None,
    emit: Callable[[str], object] = print,
) -> Iterator[Path]:
    """Back up and restore the live config with predecessor failure semantics."""
    backup_path = yaml_path.with_suffix(yaml_path.suffix + ".bak")
    copier = shutil.copy2 if copy_file is None else copy_file
    mover = shutil.move if move_file is None else move_file

    # Deliberately before the try/finally: a backup failure never attempts restore.
    copier(yaml_path, backup_path)
    emit(f"[INFO] Backup created: {backup_path.name}")

    try:
        yield backup_path
    finally:
        try:
            mover(str(backup_path), str(yaml_path))
            emit("\n[INFO] Restored original YAML from backup.")
        except Exception as error:
            emit(f"\n[WARN] Could not restore YAML from backup: {error}")
            if backup_path.exists():
                emit(f"[WARN] Backup still available at: {backup_path}")


def parse_cli_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run B1 compiler across scenarios")
    parser.add_argument(
        "--scenarios",
        default=None,
        help="Comma-separated list of scenario suffixes to run (e.g. "
        "'B_Optimised_VRE,C_Target_VRE'). When omitted, runs all "
        "scenarios discovered under A1_Outputs/.",
    )
    return parser.parse_args(argv)


def orchestrate(
    cli_args: argparse.Namespace,
    paths: B1Paths,
    *,
    scenario_discoverer: Optional[Callable[[Path], Sequence[str]]] = None,
    scenario_updater: Optional[Callable[[Path, str], None]] = None,
    compiler_runner: Optional[Callable[[Path], int]] = None,
    configuration_scope: Optional[Callable[..., object]] = None,
    emit: Callable[[str], object] = print,
) -> None:
    """Run B1 scenario compilation through explicit planning and effect seams."""
    discover = (
        list_scenario_suffixes
        if scenario_discoverer is None
        else scenario_discoverer
    )
    update = update_main_scenario if scenario_updater is None else scenario_updater
    compile_scenario = run_compiler if compiler_runner is None else compiler_runner

    def default_configuration_scope(path: Path):
        return preserved_configuration(path, emit=emit)

    config_scope = (
        default_configuration_scope
        if configuration_scope is None
        else configuration_scope
    )

    if not paths.config_path.is_file():
        emit(f"[ERROR] YAML not found: {paths.config_path}")
        sys.exit(1)
    if not paths.compiler_path.is_file():
        emit(f"[ERROR] Compiler script not found: {paths.compiler_path}")
        sys.exit(1)

    discovered = list(discover(paths.scenarios_root))
    if not discovered:
        emit("[WARN] No 'A1_Outputs_*' folders found. Nothing to do.")
        sys.exit(0)

    selection = resolve_scenarios(discovered, cli_args.scenarios)
    if selection.unknown:
        emit(
            f"[ERROR] --scenarios contains names not found under A1_Outputs/: "
            f"{list(selection.unknown)}. Discovered: {discovered}"
        )
        sys.exit(1)

    scenarios = list(selection.selected)
    if selection.filter_active:
        emit(f"[INFO] Scenario filter active: {scenarios}")

    emit(f"[INFO] Scenarios discovered: {scenarios}")

    with config_scope(paths.config_path):
        for scenario in scenarios:
            emit(f"\n[INFO] === Running scenario: {scenario} ===")
            try:
                update(paths.config_path, scenario)
                emit(
                    f"[INFO] Updated 'Main_Scenario' to '{scenario}' in "
                    f"{paths.config_path.name}"
                )
            except Exception as error:
                emit(
                    f"[ERROR] Failed to update YAML for scenario '{scenario}': "
                    f"{error}"
                )
                continue

            return_code = compile_scenario(paths.script_dir)
            if return_code != 0:
                emit(
                    f"[ERROR] B1_Compiler.py exited with code {return_code} "
                    f"for scenario '{scenario}'"
                )
            else:
                emit(
                    f"[INFO] B1_Compiler.py completed successfully for scenario "
                    f"'{scenario}'"
                )

    emit("\n[INFO] All done.")


def main() -> None:
    return orchestrate(parse_cli_args(), B1Paths.defaults())

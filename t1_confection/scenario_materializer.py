"""Materialize an exact canonical scenario selection through A3.

Root scenarios are rebuilt by the maintained A3 entrypoint.  Derived
scenarios are then rebuilt from the root declared in
``scenario_registry.json``, receive their patch layer, and finally receive any
declared direction overlay.  This module stops before B1, B2, or a solver.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime, timezone
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
from typing import Callable, Mapping, MutableMapping, Sequence

from ostram.paths import resolve_paths

try:
    from .scenario_registry import (
        ScenarioRegistry,
        load_registry,
        root_snapshots_exist,
    )
    from .sensitivity_expansion import apply_patches
except ImportError:  # direct-script execution
    from scenario_registry import (
        ScenarioRegistry,
        load_registry,
        root_snapshots_exist,
    )
    from sensitivity_expansion import apply_patches


_PROJECT_PATHS = resolve_paths()
T1_CONFECTION = _PROJECT_PATHS.legacy_runtime_root
A1_OUTPUTS = T1_CONFECTION / "A1_Outputs"
A3_ENTRYPOINT = T1_CONFECTION / "A3_process.py"
A3_PROCESS = T1_CONFECTION / "A3_process"
DEFAULT_SOASIA = A3_PROCESS / "OSTRAM_Scenario_Inputs.xlsx"
PROVENANCE_FILE = "_scenario_materialization.json"
REQUIRED_AO_FILES = (
    "A-O_Parametrization.xlsx",
    "A-O_AR_Model_Base_Year.xlsx",
    "A-O_AR_Projections.xlsx",
    "A-O_Demand.xlsx",
)


@dataclass(frozen=True)
class MaterializationPaths:
    t1_confection: Path
    a1_outputs: Path
    a3_entrypoint: Path
    a3_process: Path
    soasia: Path

    @classmethod
    def defaults(cls) -> "MaterializationPaths":
        paths = resolve_paths()
        t1_confection = paths.legacy_runtime_root
        a3_process = t1_confection / "A3_process"
        return cls(
            t1_confection=t1_confection,
            a1_outputs=t1_confection / "A1_Outputs",
            a3_entrypoint=t1_confection / "A3_process.py",
            a3_process=a3_process,
            soasia=paths.scenario_workbook,
        )


RootMaterializer = Callable[[str, Mapping[str, str]], None]
DirectionApplier = Callable[[Path, Path, int | None], dict]


def validate_control_roots(
    registry: ScenarioRegistry,
    soasia: Path,
) -> None:
    """Require the workbook Control sheet to contain only active roots."""

    module_path = soasia.parent / "_scenarios.py"
    spec = importlib.util.spec_from_file_location(
        "_ostram_materializer_scenarios",
        module_path,
    )
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot load scenario helper: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
        configs = module.read_control_sheet(soasia)
    finally:
        sys.modules.pop(spec.name, None)

    names = tuple(config.scenario for config in configs)
    if names != registry.root_names:
        raise ValueError(
            f"Control roots must be {list(registry.root_names)}, got {list(names)}"
        )
    inactive = [config.scenario for config in configs if not config.active]
    if inactive:
        raise ValueError(f"Control roots must all be active: {inactive}")


def _default_root_materializer(
    paths: MaterializationPaths,
) -> RootMaterializer:
    def materialize(root: str, environment: Mapping[str, str]) -> None:
        command = [
            sys.executable,
            "-B",
            "-m",
            "t1_confection.A3_process",
            "--scenario",
            root,
            "--soasia",
            str(paths.soasia),
        ]
        subprocess.run(
            command,
            cwd=paths.t1_confection,
            env=dict(environment),
            check=True,
        )

    return materialize


def _load_direction_module(paths: MaterializationPaths):
    module_path = (
        paths.a3_process
        / "rules_scripts"
        / "set_interconnector_direction.py"
    )
    spec = importlib.util.spec_from_file_location(
        "_ostram_direction_overlay",
        module_path,
    )
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot load direction script: {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _default_direction_applier(
    paths: MaterializationPaths,
) -> DirectionApplier:
    module = _load_direction_module(paths)

    def apply(
        target: Path,
        overlay: Path,
        study_start_year: int | None,
    ) -> dict:
        return module.run(
            target,
            skip_backup=True,
            yaml_path=overlay,
            study_start_year=study_start_year,
        )

    return apply


def materialize_scenarios(
    requested: str | Sequence[str] | None,
    *,
    paths: MaterializationPaths | None = None,
    registry: ScenarioRegistry | None = None,
    environment: MutableMapping[str, str] | None = None,
    root_materializer: RootMaterializer | None = None,
    direction_applier: DirectionApplier | None = None,
    write_provenance: bool = True,
) -> dict:
    """Materialize roots and accepted derived scenarios deterministically."""

    active_paths = MaterializationPaths.defaults() if paths is None else paths
    active_registry = (
        load_registry(active_paths.t1_confection / "scenario_registry.json")
        if registry is None
        else registry
    )
    selected = active_registry.select(requested)
    if not selected:
        raise ValueError("scenario selection is empty")

    validate_control_roots(active_registry, active_paths.soasia)
    roots = active_registry.required_roots(selected)
    if not root_snapshots_exist(active_paths.a1_outputs, roots):
        missing = [
            root
            for root in roots
            if not (
                active_paths.a1_outputs / f"_post_a2_snapshot_{root}"
            ).is_dir()
        ]
        raise FileNotFoundError(
            f"post-A2 snapshots missing for canonical roots: {missing}"
        )

    process_environment = dict(
        os.environ if environment is None else environment
    )
    resolved_dependencies = active_registry.result_dependencies(
        roots,
        t1_confection=active_paths.t1_confection,
        environment=process_environment,
    )
    process_environment.update(
        {name: str(path) for name, path in resolved_dependencies.items()}
    )

    materialize_root = (
        _default_root_materializer(active_paths)
        if root_materializer is None
        else root_materializer
    )
    apply_direction = (
        _default_direction_applier(active_paths)
        if direction_applier is None
        else direction_applier
    )

    root_records: list[dict] = []
    for root in roots:
        materialize_root(root, process_environment)
        target = active_paths.a1_outputs / f"A1_Outputs_{root}"
        missing = [
            filename
            for filename in REQUIRED_AO_FILES
            if not (target / filename).is_file()
        ]
        if missing:
            raise FileNotFoundError(
                f"A3 did not materialize required files for {root}: {missing}"
            )
        root_records.append(
            {
                "scenario": root,
                "snapshot": str(
                    active_paths.a1_outputs / f"_post_a2_snapshot_{root}"
                ),
                "target": str(target),
            }
        )

    derived_by_name = active_registry.derived_by_name
    derived_records: list[dict] = []
    for scenario in selected:
        derived = derived_by_name.get(scenario)
        if derived is None:
            continue
        log = apply_patches.build_scenario(
            scenario,
            source=derived.base_scenario,
            skip_backup=True,
            a1_outputs=active_paths.a1_outputs,
            configs=active_paths.a3_process / "rules_scripts" / "configs",
            ceiling_path=(
                active_paths.t1_confection
                / "sensitivity_expansion"
                / "reference"
                / "vre_ceilings_base.json"
            ),
            authority_path=active_paths.soasia,
        )
        target = active_paths.a1_outputs / f"A1_Outputs_{scenario}"
        direction_record: dict | None = None
        if derived.direction_overlay is not None:
            direction_log = apply_direction(
                target,
                derived.direction_overlay,
                derived.direction_study_start_year,
            )
            direction_record = {
                "overlay": str(derived.direction_overlay),
                "study_start_year": derived.direction_study_start_year,
                "skipped": bool(direction_log.get("skipped", False)),
                "projection_changes": len(
                    direction_log.get("projections", {}).get("changes", [])
                ),
                "base_year_changes": len(
                    direction_log.get("base_year", {}).get("changes", [])
                ),
            }
        derived_records.append(
            {
                "scenario": scenario,
                "base_scenario": derived.base_scenario,
                "patches": str(derived.patches),
                "patch_cells": len(log.get("cells", [])),
                "patch_rows_created": len(log.get("rows_created", [])),
                "target": str(target),
                "direction": direction_record,
            }
        )

    record = {
        "schema": "ostram-scenario-materialization-v1",
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "registry": str(active_registry.path),
        "selected_scenarios": list(selected),
        "required_roots": list(roots),
        "result_dependencies": {
            name: str(path) for name, path in resolved_dependencies.items()
        },
        "roots": root_records,
        "derived": derived_records,
        "solver_invoked": False,
    }
    if write_provenance:
        active_paths.a1_outputs.mkdir(parents=True, exist_ok=True)
        provenance_path = active_paths.a1_outputs / PROVENANCE_FILE
        provenance_path.write_text(
            json.dumps(record, indent=2),
            encoding="utf-8",
        )
        record["provenance_path"] = str(provenance_path)
    return record


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--scenarios",
        default=None,
        help=(
            "Comma-separated scenario selection. Omit for BAU plus the frozen "
            "accepted decision set."
        ),
    )
    parser.add_argument(
        "--a-result-seed",
        type=Path,
        default=None,
        help=(
            "Read-only A_Calibrated_BAU result CSV or directory used only "
            "when C_Target_VRE is selected."
        ),
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    environment = dict(os.environ)
    if args.a_result_seed is not None:
        seed = args.a_result_seed.resolve()
        if not seed.exists():
            raise FileNotFoundError(f"A result seed not found: {seed}")
        environment["OSTRAM_A_CALIBRATED_BAU_RESULT"] = str(seed)
    record = materialize_scenarios(
        args.scenarios,
        environment=environment,
    )
    print(json.dumps(record, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

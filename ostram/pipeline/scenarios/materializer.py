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
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Callable, Mapping, MutableMapping, Sequence

from ostram.paths import resolve_paths

from .registry import (
    RootScenario,
    ScenarioRegistry,
    load_registry,
    root_snapshots_exist,
)
from . import apply_patches


_PROJECT_PATHS = resolve_paths()
A1_OUTPUTS = _PROJECT_PATHS.a1_outputs
A3_ENTRYPOINT = Path(__file__).with_name("transform.py")
A3_PROCESS = Path(__file__).with_name("transformations")
DEFAULT_SOASIA = _PROJECT_PATHS.scenario_workbook
PROVENANCE_FILE = "_scenario_materialization.json"
REQUIRED_AO_FILES = (
    "A-O_Parametrization.xlsx",
    "A-O_AR_Model_Base_Year.xlsx",
    "A-O_AR_Projections.xlsx",
    "A-O_Demand.xlsx",
)


@dataclass(frozen=True)
class MaterializationPaths:
    preparation_workspace: Path
    a1_outputs: Path
    a3_entrypoint: Path
    a3_process: Path
    soasia: Path

    @classmethod
    def defaults(cls) -> "MaterializationPaths":
        paths = resolve_paths()
        a3_process = Path(__file__).with_name("transformations")
        return cls(
            preparation_workspace=paths.preparation_workspace,
            a1_outputs=paths.a1_outputs,
            a3_entrypoint=Path(__file__).with_name("transform.py"),
            a3_process=a3_process,
            soasia=paths.scenario_workbook,
        )


RootMaterializer = Callable[[str, Mapping[str, str]], None]
DirectionApplier = Callable[[Path, Path, int | None], dict]


def validate_control_roots(
    registry: ScenarioRegistry,
    soasia: Path,
) -> None:
    """Require the workbook Control sheet to contain only active roots.

    Every restriction dependency declared by the registry must also be
    declared by the matching Control ``inherit_restrictions_from`` entry, so
    the registry can never widen materialization beyond the workbook
    contract.
    """

    from .transformations import scenario_workbooks

    configs = scenario_workbooks.read_control_sheet(soasia)

    names = tuple(config.scenario for config in configs)
    if names != registry.root_names:
        raise ValueError(
            f"Control roots must be {list(registry.root_names)}, got {list(names)}"
        )
    inactive = [config.scenario for config in configs if not config.active]
    if inactive:
        raise ValueError(f"Control roots must all be active: {inactive}")
    configs_by_name = {config.scenario: config for config in configs}
    for root in registry.roots:
        undeclared = [
            prerequisite
            for prerequisite in root.restriction_dependencies
            if prerequisite
            not in configs_by_name[root.name].inherit_restrictions_from
        ]
        if undeclared:
            raise ValueError(
                f"registry restriction dependencies for {root.name} are not "
                f"declared by Control inherit_restrictions_from: {undeclared}"
            )


def _default_root_materializer(
    paths: MaterializationPaths,
    extra_arguments: Mapping[str, Sequence[str]] | None = None,
) -> RootMaterializer:
    def materialize(root: str, environment: Mapping[str, str]) -> None:
        command = [
            sys.executable,
            "-B",
            "-m",
            "ostram.pipeline.scenarios.transform",
            "--scenario",
            root,
            "--soasia",
            str(paths.soasia),
            *[str(value) for value in (extra_arguments or {}).get(root, ())],
        ]
        subprocess.run(
            command,
            cwd=resolve_paths().stage_workspace("scenarios", create=True),
            env=dict(environment),
            check=True,
        )

    return materialize


def _restriction_state_wiring(
    roots: Sequence[str],
    roots_by_name: Mapping[str, "RootScenario"],
    state_dir: Path,
) -> tuple[dict[str, list[str]], dict[str, Path]]:
    """Plan the run-state handoff for declared restriction prerequisites.

    A prerequisite exports its disposable run state (the only place generated
    Restrictions rows live); each dependent reads its prerequisites' rows
    from those exported states.  The maintained scenario workbook is never
    written.  Registries without declarations produce no wiring.
    """

    consumers = {
        root: roots_by_name[root].restriction_dependencies
        for root in roots
        if roots_by_name[root].restriction_dependencies
    }
    provider_set = {
        prerequisite for deps in consumers.values() for prerequisite in deps
    }
    providers = tuple(root for root in roots if root in provider_set)
    state_paths = {
        provider: state_dir / f"_run_state_{provider}.xlsx"
        for provider in providers
    }
    extra_arguments: dict[str, list[str]] = {}
    for provider in providers:
        extra_arguments.setdefault(provider, []).extend(
            ["--run-state-out", str(state_paths[provider])]
        )
    for consumer, prerequisites in consumers.items():
        for prerequisite in prerequisites:
            extra_arguments.setdefault(consumer, []).extend(
                [
                    "--restrictions-source",
                    f"{prerequisite}={state_paths[prerequisite]}",
                ]
            )
    return extra_arguments, state_paths


def _restriction_state_directory() -> Path:
    """Return the sole workspace-contained location for ephemeral states."""

    workspace = resolve_paths().stage_workspace("scenarios", create=True).resolve()
    state_dir = (workspace / "_restriction_states").resolve()
    try:
        state_dir.relative_to(workspace)
    except ValueError as error:
        raise ValueError(
            f"restriction state directory escapes the scenario workspace: {state_dir}"
        ) from error
    return state_dir


def _remove_restriction_state_directory(state_dir: Path) -> None:
    """Remove only the known ephemeral state directory, never an arbitrary path."""

    workspace = resolve_paths().stage_workspace("scenarios", create=True).resolve()
    resolved = state_dir.resolve()
    if resolved.name != "_restriction_states":
        raise ValueError(f"unsafe restriction state directory: {state_dir}")
    try:
        resolved.relative_to(workspace)
    except ValueError as error:
        raise ValueError(
            f"restriction state directory escapes the scenario workspace: {resolved}"
        ) from error
    shutil.rmtree(resolved, ignore_errors=True)


def _load_direction_module(paths: MaterializationPaths):
    del paths
    from .rules import set_interconnector_direction

    return set_interconnector_direction


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
        load_registry(resolve_paths().scenario_registry)
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
        execution_workspace=resolve_paths().execution_workspace,
        environment=process_environment,
    )
    process_environment.update(
        {name: str(path) for name, path in resolved_dependencies.items()}
    )

    # Declared restriction prerequisites hand their generated Restrictions
    # rows to dependents through exported disposable run states.  The wiring
    # exists only on the default transform boundary; injected materializers
    # own their own effects.
    restriction_prerequisites = {
        root: list(active_registry.roots_by_name[root].restriction_dependencies)
        for root in roots
        if active_registry.roots_by_name[root].restriction_dependencies
    }
    state_paths: dict[str, Path] = {}
    state_dir: Path | None = None
    extra_arguments: dict[str, list[str]] = {}
    if root_materializer is None and restriction_prerequisites:
        state_dir = _restriction_state_directory()
        if state_dir.exists():
            _remove_restriction_state_directory(state_dir)
        state_dir.mkdir(parents=True, exist_ok=True)
        extra_arguments, state_paths = _restriction_state_wiring(
            roots,
            active_registry.roots_by_name,
            state_dir,
        )

    materialize_root = (
        _default_root_materializer(active_paths, extra_arguments)
        if root_materializer is None
        else root_materializer
    )
    apply_direction = (
        _default_direction_applier(active_paths)
        if direction_applier is None
        else direction_applier
    )

    try:
        root_records: list[dict] = []
        for root in roots:
            materialize_root(root, process_environment)
            exported_state = state_paths.get(root)
            if (
                exported_state is not None
                and (
                    not exported_state.is_file()
                    or exported_state.stat().st_size == 0
                )
            ):
                raise FileNotFoundError(
                    f"restriction prerequisite {root} did not export a nonempty "
                    f"run state: {exported_state}"
                )
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
                configs=resolve_paths().scenario_config_root,
                ceiling_path=(
                    resolve_paths().scenario_config_root
                    / "sensitivities"
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
    finally:
        if state_dir is not None:
            _remove_restriction_state_directory(state_dir)

    record = {
        "schema": "ostram-scenario-materialization-v1",
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "registry": str(active_registry.path),
        "selected_scenarios": list(selected),
        "required_roots": list(roots),
        "restriction_prerequisites": restriction_prerequisites,
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

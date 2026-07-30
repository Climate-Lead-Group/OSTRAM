"""Behavior-preserving planning and effect sequencing for Stage A3.

The public command remains ``t1_confection/A3_process.py``.  This module owns
only orchestration: path/scenario planning, snapshot restoration, ordered stage
dispatch, delivery, and the predecessor cleanup boundary.  Workbook and model
transformations remain in their existing helpers and are supplied explicitly
through :class:`A3Dependencies`.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, MutableMapping, Sequence


PWR_MIN_PIN_ROOT_SCENARIOS = frozenset(
    {"A_Calibrated_BAU", "B_Optimised_VRE", "C_Target_VRE"}
)


@dataclass(frozen=True)
class A3Paths:
    """Script-anchored filesystem roots used by one A3 invocation."""

    t1_confection: Path
    process_dir: Path
    default_soasia: Path

    @classmethod
    def from_entrypoint(cls, entrypoint: str | Path) -> "A3Paths":
        t1_confection = Path(entrypoint).resolve().parent
        process_dir = t1_confection / "A3_process"
        return cls(
            t1_confection=t1_confection,
            process_dir=process_dir,
            default_soasia=process_dir / "SOASIA_OSeMOSYS_Template_v18.xlsx",
        )


@dataclass(frozen=True)
class A3Plan:
    """Resolved, side-effect-free plan for one direct A3 invocation."""

    scenario: str
    rules_scripts: tuple[str, ...]
    inherit_from: tuple[str, ...]
    soasia: Path
    input_dir: Path
    output_dir: Path
    snapshot_dir: Path
    workdir_base: Path
    keep_workdir: bool


@dataclass(frozen=True)
class A3Dependencies:
    """Existing A3 operations and narrow effect seams used by orchestration."""

    resolve_scenario_config: Callable[
        [argparse.Namespace, Path], tuple[str, list[str], list[str]]
    ]
    resolve_path: Callable[[Path | str], Path]
    build_workdir: Callable[[Path, str, list[str], str], dict[str, Path]]
    materialize_scenario_template: Callable[[Path, str, Path], object]
    stage_0_5_rnwbio: Callable[[Path, Path], object]
    stage_1_scripts_1_to_5: Callable[[Path], object]
    stage_1b: Callable[[Path, Path, Path], object]
    stage_2_and_2_5: Callable[[Path, Path, Path], object]
    stage_3_fix_2: Callable[[Path, Path], Path]
    stage_4_consolidate: Callable[[Path, Path, Path, Path], object]
    stage_4_5_apply_inherited_restrictions: Callable[
        [Path, Path, list[str]], object
    ]
    stage_5_rules_scripts: Callable[[Path, Path, list[str]], object]
    stage_ws3_interconnector_costs: Callable[
        [Path, Path, Path | None], object
    ]
    stage_ws3_internal_transmission: Callable[[Path], object]
    stage_ws3_internal_tx_losses: Callable[[Path], object]
    stage_ws4_pwr_min_pin: Callable[[Path, str], object]
    stage_6_sync_og_to_ts20: Callable[[Path, Path], object]
    stage_6_persist_restrictions: Callable[
        [Path, Path, str, list[str]], object
    ]
    deliver_outputs: Callable[[Path, Path], object]
    remove_tree: Callable[..., object]
    copy_tree: Callable[..., object]
    copy_file: Callable[..., object]
    environment: MutableMapping[str, str]
    clock: Callable[[], float]
    timestamp_now: Callable[[], object]
    banner: Callable[[str], object]
    emit: Callable[[str], object]


def resolve_plan(
    cli_args: argparse.Namespace,
    paths: A3Paths,
    dependencies: A3Dependencies,
) -> A3Plan:
    """Resolve scenario, dependency, and destination choices without effects."""
    if not paths.process_dir.is_dir():
        raise SystemExit(f"ERROR: A3_process folder missing: {paths.process_dir}")

    soasia = (
        cli_args.soasia
        if cli_args.soasia is not None
        else paths.default_soasia
    )
    scenario, rules_scripts, inherit_from = (
        dependencies.resolve_scenario_config(cli_args, soasia)
    )

    if cli_args.input_dir is not None:
        input_dir = dependencies.resolve_path(cli_args.input_dir)
    else:
        input_dir = (
            paths.t1_confection
            / "A1_Outputs"
            / f"A1_Outputs_{scenario}"
        )
    output_dir = (
        dependencies.resolve_path(cli_args.output_dir)
        if cli_args.output_dir is not None
        else input_dir
    )

    return A3Plan(
        scenario=scenario,
        rules_scripts=tuple(rules_scripts),
        inherit_from=tuple(inherit_from),
        soasia=soasia,
        input_dir=input_dir,
        output_dir=output_dir,
        snapshot_dir=(
            paths.t1_confection
            / "A1_Outputs"
            / f"_post_a2_snapshot_{scenario}"
        ),
        workdir_base=paths.process_dir,
        keep_workdir=cli_args.keep_workdir,
    )


def execute_plan(
    plan: A3Plan,
    dependencies: A3Dependencies,
    input_files: Sequence[str],
) -> int:
    """Execute the frozen predecessor stage order through injected boundaries."""
    if not plan.snapshot_dir.is_dir():
        raise SystemExit(
            f"ERROR: snapshot post-A2 not found: {plan.snapshot_dir}\n"
            f"       Run A1 + A2 (for BAU) first; A2 creates the snapshot."
        )

    t_start = dependencies.clock()
    timestamp = dependencies.timestamp_now().strftime("%Y%m%d_%H%M%S")
    dependencies.banner(f"A3 workflow run @ {timestamp}")
    dependencies.emit(f"  scenario          : {plan.scenario}")
    dependencies.emit(f"  input-dir         : {plan.input_dir}")
    dependencies.emit(f"  output-dir        : {plan.output_dir}")
    dependencies.emit(f"  snapshot (source) : {plan.snapshot_dir}")
    dependencies.emit(f"  SOASIA v18        : {plan.soasia}")
    dependencies.emit(
        f"  rules_scripts     : {list(plan.rules_scripts) or '(none)'}"
    )
    dependencies.emit(
        f"  inherit_from      : {list(plan.inherit_from) or '(none)'}"
    )

    if plan.input_dir.exists():
        dependencies.remove_tree(plan.input_dir)
    dependencies.copy_tree(plan.snapshot_dir, plan.input_dir)
    dependencies.emit(
        f"  -> {plan.input_dir.name} restored from {plan.snapshot_dir.name}"
    )

    rules_scripts = list(plan.rules_scripts)
    inherit_from = list(plan.inherit_from)
    runtime_paths = dependencies.build_workdir(
        plan.workdir_base,
        timestamp,
        rules_scripts,
        plan.scenario,
    )
    workdir = runtime_paths["wd"]
    stage1 = runtime_paths["s1"]
    stage1b = runtime_paths["s1b"]
    stage2 = runtime_paths["s2"]
    stage3 = runtime_paths["s3"]
    stage5 = runtime_paths["s5"]
    dependencies.emit(f"  workdir           : {workdir}")

    materialized_template: Path | None = None
    if plan.soasia.is_file():
        materialized_template = (
            workdir / f"_materialized_{plan.scenario}.xlsx"
        )
        dependencies.banner(
            f"Stage 0 — materialize scenario template for '{plan.scenario}'"
        )
        dependencies.materialize_scenario_template(
            plan.soasia,
            plan.scenario,
            materialized_template,
        )
        dependencies.environment["OSTRAM_TEMPLATE_PATH"] = str(
            materialized_template
        )
        dependencies.emit(
            f"    materialized -> {materialized_template.name}"
        )
        dependencies.emit(
            "    OSTRAM_TEMPLATE_PATH set; stage 1 will read it instead of v17"
        )

    for filename in input_files:
        source = plan.input_dir / filename
        if not source.exists():
            raise SystemExit(f"ERROR: input file missing: {source}")
        dependencies.copy_file(source, stage1 / filename)

    dependencies.stage_0_5_rnwbio(workdir, stage1)
    dependencies.stage_1_scripts_1_to_5(stage1)
    dependencies.stage_1b(workdir, stage1, stage1b)
    dependencies.stage_2_and_2_5(workdir, stage1b, stage2)
    parameter_for_stage4 = dependencies.stage_3_fix_2(stage2, stage3)
    dependencies.stage_4_consolidate(
        stage1,
        stage3,
        stage5,
        parameter_for_stage4,
    )
    dependencies.stage_4_5_apply_inherited_restrictions(
        stage5,
        plan.soasia,
        inherit_from,
    )
    dependencies.stage_5_rules_scripts(
        workdir,
        stage5,
        rules_scripts,
    )
    dependencies.stage_ws3_interconnector_costs(
        stage5,
        plan.soasia,
        materialized_template,
    )
    dependencies.stage_ws3_internal_transmission(stage5)
    dependencies.stage_ws3_internal_tx_losses(stage5)
    if plan.scenario in PWR_MIN_PIN_ROOT_SCENARIOS:
        dependencies.stage_ws4_pwr_min_pin(stage5, plan.scenario)
    dependencies.stage_6_sync_og_to_ts20(workdir, stage1)
    dependencies.stage_6_persist_restrictions(
        stage5,
        plan.soasia,
        plan.scenario,
        rules_scripts,
    )
    dependencies.deliver_outputs(stage5, plan.output_dir)

    if not plan.keep_workdir:
        dependencies.remove_tree(workdir, ignore_errors=True)
        dependencies.emit(f"\n  Cleaned up workdir: {workdir.name}")
    else:
        dependencies.emit(f"\n  Workdir preserved: {workdir}")
    dependencies.environment.pop("OSTRAM_TEMPLATE_PATH", None)

    elapsed = dependencies.clock() - t_start
    dependencies.banner(f"DONE in {elapsed:.1f}s")
    return 0


def orchestrate_a3(
    cli_args: argparse.Namespace,
    paths: A3Paths,
    dependencies: A3Dependencies,
    input_files: Sequence[str],
) -> int:
    """Plan and execute one A3 invocation without hidden dependency lookup."""
    plan = resolve_plan(cli_args, paths, dependencies)
    return execute_plan(plan, dependencies, input_files)

"""Explicit orchestration and solver boundaries for the B2 workflow.

The computational transformations remain in :mod:`B2_Executing_OG_Model` and
are supplied here as dependencies.  Keeping this module dependency-injected
makes the orchestration testable without making model, otoole, or solver calls.
"""

from __future__ import annotations

import argparse
import os
import shutil
import time
import traceback
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Callable

from ostram.paths import resolve_paths
from ostram.validation.profile import validate_active_compiled_domain


@dataclass(frozen=True)
class B2Dependencies:
    """Existing B2 operations used by the top-level orchestration."""

    process_scenario_folder: Callable[..., Any]
    run_otoole_conversion: Callable[..., bool]
    run_preprocessing_script: Callable[..., Any]
    run_days_in_day_type_patcher: Callable[..., Any]
    run_storage_delay_patcher: Callable[..., Any]
    run_strip_storage_patcher: Callable[..., Any]
    run_open_pwrbck_patcher: Callable[..., Any]
    run_reserve_margin_repair_patcher: Callable[..., Any]
    run_reserve_margin_xlsx_patcher: Callable[..., Any]
    generate_combined_input_file: Callable[..., Any]
    export_root_datafile: Callable[..., Any]
    main_executer: Callable[..., Any]
    chunk_scenarios: Callable[..., Any]
    delete_files: Callable[..., Any]
    concatenate_all_scenarios: Callable[..., Any]
    load_annualizer: Callable[[], Callable[..., Any]]
    yaml_safe_load: Callable[..., Any]
    mp_module: Any


@dataclass(frozen=True)
class B2RunPlan:
    """Resolved configuration, scenarios, and paths for one B2 run."""

    here: Path
    params: dict[str, Any]
    params_a2: dict[str, Any]
    scenarios: list[str]
    main_scenario_name: str
    base_input_path: str
    template_path: str
    base_output_path: str
    compile_only: bool = False


def build_cli_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Execute OSeMOSYS model across scenarios"
    )
    parser.add_argument(
        "--scenarios",
        default=None,
        help=(
            "Comma-separated list of scenario names to run (e.g. "
            "'B_Optimised_VRE,C_Target_VRE'). When omitted, runs all "
            "scenarios found in the A2 output directory."
        ),
    )
    parser.add_argument(
        "--compile-only",
        action="store_true",
        help=(
            "Generate the final preprocessed and patched text inputs, then "
            "stop before matrix creation, every solver adapter, cleanup, and "
            "result post-processing."
        ),
    )
    return parser


def parse_arguments(argv: list[str] | None = None) -> argparse.Namespace:
    return build_cli_parser().parse_args(argv)


def resolve_here() -> Path:
    """Return the explicit mutable execution workspace."""

    return resolve_paths().stage_workspace("execution", create=True)


def apply_configuration_overrides(params: dict[str, Any]) -> None:
    """Apply the existing storage-delay precedence mutations in place."""

    if params.get("storage_delay_active", False):
        if params.get("strip_storage_active", False):
            print(
                "[storage_delay] strip_storage_active forced to False "
                "(mutually exclusive)"
            )
            params["strip_storage_active"] = False
        params.setdefault("storage_delay_model_input", params["osemosys_model"])
        params.setdefault(
            "storage_delay_model_output",
            "osemosys_fast_preprocessed_storage_delay.txt",
        )
        params["osemosys_model"] = params["storage_delay_model_output"]
        if params.get("storage_delay_prefix_final_files"):
            params["prefix_final_files"] = params[
                "storage_delay_prefix_final_files"
            ]
        print(f"[storage_delay] osemosys_model -> {params['osemosys_model']}")
        print(
            "[storage_delay] prefix_final_files -> "
            f"{params['prefix_final_files']}"
        )


def apply_compile_only_overrides(params: dict[str, Any]) -> None:
    """Disable every matrix, solver, cleanup, and result boundary."""

    params["execute_model"] = False
    params["create_matrix"] = False
    params["reuse_existing_sol"] = False
    params["concat_otoole_csv"] = False
    params["concat_scenarios_csv"] = False
    params["annualize_capital"] = False
    params["del_files"] = False


def resolve_scenarios(
    base_input_path: str,
    params: dict[str, Any],
    params_a2: dict[str, Any],
    scenario_filter: str | None,
) -> list[str]:
    """Discover and filter scenarios with the predecessor's exact semantics."""

    scenarios = sorted(os.listdir(base_input_path))
    try:
        scenarios.remove("Default")
    except ValueError:
        pass

    if params["only_main_scenario"]:
        scenarios = []
        scenarios.append(params_a2["xtra_scen"]["Main_Scenario"])

    if scenario_filter:
        requested = [
            scenario.strip()
            for scenario in scenario_filter.split(",")
            if scenario.strip()
        ]
        requested_unique = list(dict.fromkeys(requested))
        discovered = set(scenarios)
        unknown = [scenario for scenario in requested if scenario not in discovered]
        if unknown:
            print(
                f"[ERROR] --scenarios contains names not found in "
                f"{base_input_path}: {unknown}. Discovered: {scenarios}"
            )
            raise SystemExit(1)
        scenarios = [
            scenario for scenario in requested_unique if scenario in discovered
        ]
        print(f"[INFO] Scenario filter active: {scenarios}")

    return scenarios


def build_run_plan(
    here: Path,
    scenario_filter: str | None,
    *,
    yaml_safe_load: Callable[..., Any],
    compile_only: bool = False,
) -> B2RunPlan:
    """Load configuration and resolve the complete non-executing B2 plan."""

    project = resolve_paths()
    local_contract = (here / "Config_MOMF_T1_AB.yaml").is_file()
    execution_config = (
        here / "Config_MOMF_T1_AB.yaml"
        if local_contract
        else project.execution_config
    )
    compilation_config = (
        here / "Config_MOMF_T1_A.yaml"
        if local_contract
        else project.compilation_config
    )

    with execution_config.open(
        "r", encoding="utf-8"
    ) as config_file:
        params = yaml_safe_load(config_file)

    if local_contract:
        for key in (
            "A2_output",
            "A2_output_otoole",
            "Miscellaneous",
            "executables",
            "osemosys_model",
            "reserve_margin_xlsx_workbook",
        ):
            value = params.get(key)
            if value:
                path = Path(str(value)).expanduser()
                params[key] = str(
                    path.resolve() if path.is_absolute() else (here / path).resolve()
                )
    else:
        params.update(
            {
                "A2_output": str(project.compiled_parameters),
                "A2_output_otoole": str(project.otoole_outputs),
                "Miscellaneous": str(project.compilation_resources),
                "templates": str(project.compilation_resources / "templates"),
                "otoole_config": str(
                    project.compilation_resources / "conversion_format.yaml"
                ),
                "conv_format": str(
                    project.compilation_resources / "conversion_format.yaml"
                ),
                "executables": str(project.executables),
                "outputs": str(project.outputs),
                "osemosys_model": str(project.maintained_model),
                "preprocess_data": str(Path(__file__).with_name("preprocess.py")),
                "reserve_margin_xlsx_workbook": str(
                    project.execution_inputs / "firm_capacity_fallbacks_by_cr.xlsx"
                ),
            }
        )

        storage_delay_output = params.get("storage_delay_model_output")
        if storage_delay_output:
            params["storage_delay_model_output"] = Path(
                str(storage_delay_output)
            ).name

    # Storage-delay defaults must be derived after profile model/path
    # authorities have replaced any declarative manifest tokens.
    apply_configuration_overrides(params)
    if compile_only:
        apply_compile_only_overrides(params)

    with compilation_config.open(
        "r", encoding="utf-8"
    ) as config_file:
        params_a2 = yaml_safe_load(config_file)

    base_input_path = os.path.join(here, params["A2_output"])
    template_path = os.path.join(
        here, params["Miscellaneous"], params["templates"]
    )
    base_output_path = os.path.join(here, params["A2_output_otoole"])
    scenarios = resolve_scenarios(
        base_input_path,
        params,
        params_a2,
        scenario_filter,
    )

    return B2RunPlan(
        here=here,
        params=params,
        params_a2=params_a2,
        scenarios=scenarios,
        main_scenario_name=params_a2["xtra_scen"]["Main_Scenario"],
        base_input_path=base_input_path,
        template_path=template_path,
        base_output_path=base_output_path,
        compile_only=compile_only,
    )


def run_compiled_input_stage(
    plan: B2RunPlan,
    dependencies: B2Dependencies,
) -> None:
    """Generate the preprocessed, patched, solver-consumed text inputs."""

    params = plan.params
    for scenario_name in plan.scenarios:
        if params["A2_otoole_outputs"]:
            dependencies.process_scenario_folder(
                base_input_path=plan.base_input_path,
                template_path=plan.template_path,
                base_output_path=plan.base_output_path,
                scenario_name=scenario_name,
            )
        domain = validate_active_compiled_domain(
            Path(plan.base_output_path) / scenario_name
        )
        if domain is not None:
            counts = domain["compiled"]
            print(
                "[profile-domain] compiled domain accepted: "
                f"TECHNOLOGY={counts['TECHNOLOGY']['count']}, "
                f"FUEL={counts['FUEL']['count']}, "
                f"scenario={scenario_name}"
            )
        if params["write_txt_model"]:
            conversion_ok = dependencies.run_otoole_conversion(
                base_output_path=plan.base_output_path,
                scenario_name=scenario_name,
                params=params,
            )

            if conversion_ok:
                dependencies.run_preprocessing_script(params, scenario_name)
                dependencies.run_days_in_day_type_patcher(params, scenario_name)
                dependencies.run_storage_delay_patcher(params, scenario_name)
                dependencies.run_strip_storage_patcher(params, scenario_name)
                dependencies.run_open_pwrbck_patcher(params, scenario_name)
                dependencies.run_reserve_margin_repair_patcher(
                    params, scenario_name
                )
                dependencies.run_reserve_margin_xlsx_patcher(
                    params, scenario_name
                )
            else:
                message = (
                    f"otoole conversion failed for scenario '{scenario_name}'; "
                    "B2 cannot report overall success"
                )
                print(f"[FAILED] {message}")
                raise RuntimeError(message)

        input_folder = os.path.join(
            plan.here, plan.base_output_path, scenario_name
        )
        output_folder = os.path.join(
            plan.here, params["executables"], scenario_name + "_0"
        )
        os.makedirs(input_folder, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)
        dependencies.generate_combined_input_file(
            input_folder,
            output_folder,
            scenario_name + "_0",
        )

    if params["write_txt_model"] and plan.main_scenario_name in plan.scenarios:
        dependencies.export_root_datafile(
            plan.here,
            params,
            plan.main_scenario_name,
        )


def run_execution_stage(
    plan: B2RunPlan,
    dependencies: B2Dependencies,
) -> None:
    """Dispatch the unchanged parallel or linear per-scenario execution route."""

    params = plan.params
    if params["execute_model"] or params["create_matrix"]:
        if params["parallel"]:
            print("Started parallelization of model execution")
            max_x_per_iter = params["max_x_per_iter"]
            scenario_chunks = dependencies.chunk_scenarios(
                plan.scenarios,
                max_x_per_iter,
            )
            for scenario_chunk in scenario_chunks:
                processes = []
                for scenario_name in scenario_chunk:
                    process = dependencies.mp_module.Process(
                        target=dependencies.main_executer,
                        args=(params, scenario_name, plan.here),
                    )
                    processes.append((scenario_name, process))
                    process.start()
                for _scenario_name, process in processes:
                    process.join()
                failures = [
                    (scenario_name, process.exitcode)
                    for scenario_name, process in processes
                    if process.exitcode not in (0, None)
                ]
                if failures:
                    detail = ", ".join(
                        f"{scenario} (exit {exitcode})"
                        for scenario, exitcode in failures
                    )
                    print(f"[FAILED] Parallel B2 worker failure: {detail}")
                    raise SystemExit(failures[0][1])
        else:
            print("Started linear executions")
            for scenario_name in plan.scenarios:
                dependencies.main_executer(params, scenario_name, plan.here)


def run_cleanup_stage(
    plan: B2RunPlan,
    dependencies: B2Dependencies,
) -> None:
    """Apply the predecessor's optional intermediate-file cleanup."""

    params = plan.params
    for scenario_name in plan.scenarios:
        if params["del_files"]:
            folder_scenario = os.path.join(
                plan.here,
                params["executables"],
                scenario_name + "_0",
            )
            outputs_otoole_csvs = os.path.join(
                plan.here,
                folder_scenario,
                params["outputs"],
            )
            data_file = os.path.join(
                plan.here,
                folder_scenario,
                scenario_name + "_0" + ".txt",
            )
            sol_file = os.path.join(
                plan.here,
                folder_scenario,
                params["preprocess_data_name"]
                + scenario_name
                + "_0"
                + params["output_files"]
                + ".sol",
            )
            if os.path.exists(outputs_otoole_csvs):
                shutil.rmtree(outputs_otoole_csvs)

            if params["solver"] in ["glpk", "cbc", "cplex"]:
                dependencies.delete_files(
                    sol_file,
                    data_file,
                    params["solver"],
                )

            print(
                f"✅ Intermediate files for scenario {scenario_name}_0 "
                "deleted successfully."
            )
            print(
                "\n#---------------------------------------------------------------"
                "---------------#"
            )


def run_final_postprocessing_stage(
    plan: B2RunPlan,
    dependencies: B2Dependencies,
) -> None:
    """Run cross-scenario concatenation and optional annualization."""

    params = plan.params
    if params["concat_scenarios_csv"]:
        input_output_path, output_output_path, combined_output_path = (
            dependencies.concatenate_all_scenarios(plan.here, params)
        )
        print("✅ Inputs and outputs concatenated for all scenarios successfully.")
        print(
            f"The files are: ({input_output_path}), ({output_output_path}) "
            f"and ({combined_output_path})"
        )

    if params.get("annualize_capital", False):
        try:
            print("\n")
            print("#" * 80)
            print("# CAPITAL INVESTMENT ANNUALIZATION")
            print("#" * 80)

            annualize_capital_investment = dependencies.load_annualizer()
            combined_file_path = os.path.join(
                plan.here,
                params["prefix_final_files"]
                + "Combined_Inputs_Outputs.csv",
            )
            if os.path.exists(combined_file_path):
                print(f"Starting annualization for: {combined_file_path}")
                annualize_capital_investment(
                    input_file_path=combined_file_path,
                    verbose=True,
                )
                print("✅ Capital investment annualization completed successfully.")

                today = date.today().isoformat()
                dated_combined = combined_file_path.replace(
                    ".csv", f"_{today}.csv"
                )
                shutil.copy2(combined_file_path, dated_combined)
                print(f"✅ Annualized file copied to: {dated_combined}")
                print("#" * 80)
            else:
                print(
                    f"⚠️  WARNING: Combined file not found at "
                    f"{combined_file_path}"
                )
                print("Skipping capital investment annualization.")
                print("#" * 80)
        except Exception as error:
            print(f"❌ ERROR during capital investment annualization: {error}")
            print("Continuing without annualization...")
            traceback.print_exc()
            print("#" * 80)
    else:
        combined_file_path = os.path.join(
            plan.here,
            params["prefix_final_files"] + "Combined_Inputs_Outputs.csv",
        )
        if os.path.exists(combined_file_path):
            today = date.today().isoformat()
            dated_combined = combined_file_path.replace(
                ".csv", f"_{today}.csv"
            )
            shutil.copy2(combined_file_path, dated_combined)
            print(f"✅ Combined file copied to: {dated_combined}")


def orchestrate_b2(
    dependency_factory: Callable[[], B2Dependencies],
    *,
    argv: list[str] | None = None,
    set_here: Callable[[Path], None],
) -> None:
    """Top-level B2 orchestration with no work performed at import time."""

    start1 = time.time()
    cli_args = parse_arguments(argv)

    here = resolve_here()
    set_here(here)

    dependencies = dependency_factory()
    plan = build_run_plan(
        here,
        cli_args.scenarios,
        yaml_safe_load=dependencies.yaml_safe_load,
        compile_only=cli_args.compile_only,
    )
    run_compiled_input_stage(plan, dependencies)
    if plan.compile_only:
        print(
            "[INFO] Compile-only gate complete; matrix, solver, cleanup, "
            "and result stages were not invoked."
        )
        return
    run_execution_stage(plan, dependencies)
    run_cleanup_stage(plan, dependencies)

    end_1 = time.time()
    time_elapsed_1 = -start1 + end_1
    print(
        str(time_elapsed_1) + " seconds /",
        str(time_elapsed_1 / 60) + " minutes",
    )

    start2 = time.time()
    run_final_postprocessing_stage(plan, dependencies)

    end_2 = time.time()
    time_elapsed_2 = -start2 + end_2
    print(
        str(time_elapsed_2) + " seconds /",
        str(time_elapsed_2 / 60) + " minutes",
    )
    print(
        "\n#------------------------------------------------------------------"
        "------------#"
    )

    time_elapsed_3 = -start1 + end_2
    print(
        str(time_elapsed_3) + " seconds /",
        str(time_elapsed_3 / 60) + " minutes",
    )
    print("*: For all effects, we have finished the work of this script.")


@dataclass(frozen=True)
class ScenarioExecutionDependencies:
    """Side effects used by one scenario's matrix/solver/output boundary."""

    run_process: Callable[..., Any]
    check_environment: Callable[[str], Any]
    get_executable: Callable[[str], str]
    path_exists: Callable[[str], bool]
    remove_file: Callable[[str], Any]
    python_executable: str


@dataclass(frozen=True)
class ScenarioExecutionPaths:
    folder_scenario: str
    data_file: str
    output_file: str


def resolve_scenario_execution_paths(
    params: dict[str, Any],
    scenario_name: str,
    here: Any,
) -> ScenarioExecutionPaths:
    """Resolve the active patched data and output bases, preserving redirects."""

    folder_scenario = os.path.join(
        here,
        params["executables"],
        scenario_name + "_0",
    )
    data_file = os.path.join(
        folder_scenario,
        params["preprocess_data_name"] + scenario_name + "_0",
    )
    output_file = os.path.join(
        folder_scenario,
        params["preprocess_data_name"]
        + scenario_name
        + "_0"
        + params["output_files"],
    )

    if params.get("storage_delay_active", False):
        storage_delay_suffix = params.get(
            "storage_delay_suffix", "StorageDelayN5"
        )
        base = params["preprocess_data_name"] + scenario_name + "_0"
        data_file = os.path.join(
            folder_scenario,
            f"{base}_{storage_delay_suffix}",
        )
        output_file = os.path.join(
            folder_scenario,
            f"{base}_{storage_delay_suffix}{params['output_files']}",
        )
        print(f"[storage_delay] redirecting solver to: {data_file}.txt")

    if params.get("strip_storage_active", False):
        strip_suffix = params.get("strip_storage_suffix", "NoStorage")
        base = params["preprocess_data_name"] + scenario_name + "_0"
        data_file = os.path.join(folder_scenario, f"{base}_{strip_suffix}")
        output_file = os.path.join(
            folder_scenario,
            f"{base}_{strip_suffix}{params['output_files']}",
        )
        print(f"[strip_storage] redirecting solver to: {data_file}.txt")

    if params.get("open_pwrbck_active", False):
        open_pwrbck_suffix = params.get("open_pwrbck_suffix", "OpenBCK")
        base = params["preprocess_data_name"] + scenario_name + "_0"
        chain_parts = []
        if params.get("storage_delay_active", False):
            chain_parts.append(
                params.get("storage_delay_suffix", "StorageDelayN5")
            )
        if params.get("strip_storage_active", False):
            chain_parts.append(params.get("strip_storage_suffix", "NoStorage"))
        chain_parts.append(open_pwrbck_suffix)
        chain = "_".join(chain_parts)
        data_file = os.path.join(folder_scenario, f"{base}_{chain}")
        output_file = os.path.join(
            folder_scenario,
            f"{base}_{chain}{params['output_files']}",
        )
        print(f"[open_pwrbck] redirecting solver to: {data_file}.txt")

    if params.get("reserve_margin_repair_active", False):
        reserve_margin_suffix = params.get(
            "reserve_margin_repair_suffix", "RMRepair"
        )
        base = params["preprocess_data_name"] + scenario_name + "_0"
        chain_parts = []
        if params.get("storage_delay_active", False):
            chain_parts.append(
                params.get("storage_delay_suffix", "StorageDelayN5")
            )
        if params.get("strip_storage_active", False):
            chain_parts.append(params.get("strip_storage_suffix", "NoStorage"))
        if params.get("open_pwrbck_active", False):
            chain_parts.append(params.get("open_pwrbck_suffix", "OpenBCK"))
        chain_parts.append(reserve_margin_suffix)
        chain = "_".join(chain_parts)
        data_file = os.path.join(folder_scenario, f"{base}_{chain}")
        output_file = os.path.join(
            folder_scenario,
            f"{base}_{chain}{params['output_files']}",
        )
        print(f"[reserve_margin_repair] redirecting solver to: {data_file}.txt")

    if params.get("reserve_margin_xlsx_active", False):
        reserve_margin_xlsx_suffix = params.get(
            "reserve_margin_xlsx_suffix", "RMCarefulXLSX"
        )
        base = params["preprocess_data_name"] + scenario_name + "_0"
        chain_parts = []
        if params.get("storage_delay_active", False):
            chain_parts.append(
                params.get("storage_delay_suffix", "StorageDelayN5")
            )
        if params.get("strip_storage_active", False):
            chain_parts.append(params.get("strip_storage_suffix", "NoStorage"))
        if params.get("open_pwrbck_active", False):
            chain_parts.append(params.get("open_pwrbck_suffix", "OpenBCK"))
        if params.get("reserve_margin_repair_active", False):
            chain_parts.append(
                params.get("reserve_margin_repair_suffix", "RMRepair")
            )
        chain_parts.append(reserve_margin_xlsx_suffix)
        chain = "_".join(chain_parts)
        data_file = os.path.join(folder_scenario, f"{base}_{chain}")
        output_file = os.path.join(
            folder_scenario,
            f"{base}_{chain}{params['output_files']}",
        )
        print(f"[reserve_margin_xlsx] redirecting solver to: {data_file}.txt")

    return ScenarioExecutionPaths(
        folder_scenario=folder_scenario,
        data_file=data_file,
        output_file=output_file,
    )


def build_matrix_command(
    params: dict[str, Any],
    solver: str,
    paths: ScenarioExecutionPaths,
    reuse_solution: bool,
) -> list[str] | None:
    """Plan GLPSOL matrix preparation without invoking a process."""

    if solver != "glpk" and params["create_matrix"] and not reuse_solution:
        return [
            "glpsol", "-m", str(params["osemosys_model"]),
            "-d", f"{paths.data_file}.txt",
            "--wlp", f"{paths.output_file}.lp", "--check",
        ]
    return None


def run_matrix_preparation(
    command: list[str],
    process_runner: Callable[..., Any],
) -> None:
    """The distinct, injectable matrix process boundary."""

    process_runner(
        command,
        cwd=str(resolve_paths().stage_workspace("execution", create=True)),
        check=True,
    )


def invoke_solver_command(
    command: list[str],
    process_runner: Callable[..., Any],
) -> None:
    """The named external solver invocation boundary."""

    process_runner(
        command,
        cwd=str(resolve_paths().stage_workspace("execution", create=True)),
        check=True,
    )


def validate_cbc_solution(solution_file: str | os.PathLike[str]) -> str:
    """Require CBC's solution header to declare an optimal solution.

    CBC returns process status zero for model-level outcomes such as an
    infeasible linear relaxation.  The solution header is therefore the
    authoritative status boundary; result conversion must never turn an
    infeasible incumbent into a successful OSTRAM report.
    """

    path = Path(solution_file)
    if not path.is_file():
        raise FileNotFoundError(f"CBC solution file not found: {path}")
    with path.open("r", encoding="utf-8", errors="replace") as stream:
        status = next((line.strip() for line in stream if line.strip()), "")
    if not status.lower().startswith("optimal - objective value"):
        raise RuntimeError(
            f"CBC did not produce an optimal solution: {status or '<empty status>'} "
            f"({path})"
        )
    return status


class SolverAdapter:
    """Prepare and invoke supported solver commands behind one explicit seam."""

    def __init__(self, dependencies: ScenarioExecutionDependencies) -> None:
        self.dependencies = dependencies

    def prepare_command(
        self,
        solver: str,
        params: dict[str, Any],
        paths: ScenarioExecutionPaths,
    ) -> list[str] | None:
        """Preserve solver-specific preparation order before matrix execution."""

        if solver == "glpk":
            self.dependencies.check_environment("glpsol")
            return [
                "glpsol", "-m", str(params["osemosys_model"]),
                "-d", f"{paths.data_file}.txt",
                "--wglp", f"{paths.output_file}.glp",
                "--write", f"{paths.output_file}.sol",
            ]

        if solver == "cbc":
            solution_file = paths.output_file + ".sol"
            if self.dependencies.path_exists(solution_file):
                self.dependencies.remove_file(solution_file)
            self.dependencies.check_environment("cbc")
            cbc_random_seed = params.get("cbc_random_seed", 12345)
            return [
                "cbc", f"{paths.output_file}.lp",
                "randomSeed", str(cbc_random_seed),
                "randomCbcSeed", str(cbc_random_seed),
                "-seconds", str(params["iteration_time"]),
                "solve", "-solu", f"{paths.output_file}.sol",
            ]

        if solver == "cplex":
            for solution_file in (
                paths.output_file + ".sol",
                paths.output_file + ".feasopt.sol",
            ):
                if self.dependencies.path_exists(solution_file):
                    self.dependencies.remove_file(solution_file)
            cplex_threads = params["cplex_threads"]
            cplex_random_seed = params.get("cplex_random_seed", 12345)
            self.dependencies.check_environment("cplex")
            return [
                "cplex", "-c",
                f"set logfile {paths.output_file}.cplex.log",
                f"read {paths.output_file}.lp",
                f"set threads {cplex_threads}",
                f"set randomseed {cplex_random_seed}",
                "set parallel 1", "optimize",
                f"write {paths.output_file}.sol",
            ]

        if solver == "gurobi":
            solution_file = paths.output_file + ".sol"
            if self.dependencies.path_exists(solution_file):
                self.dependencies.remove_file(solution_file)
            gurobi_threads = params["gurobi_threads"]
            gurobi_seed = params.get("gurobi_seed", 12345)
            self.dependencies.check_environment("gurobi_cl")
            return [
                "gurobi_cl", f"Threads={gurobi_threads}",
                f"Seed={gurobi_seed}",
                f"ResultFile={paths.output_file}.sol",
                f"{paths.output_file}.lp",
            ]

        return None

    def invoke(self, command: list[str]) -> None:
        invoke_solver_command(command, self.dependencies.run_process)


def run_scenario_output_stage(
    params: dict[str, Any],
    scenario_name: str,
    here: Any,
    solver: str,
    paths: ScenarioExecutionPaths,
    dependencies: ScenarioExecutionDependencies,
) -> None:
    """Convert and concatenate one scenario's solver outputs."""

    if params["execute_model"]:
        print(f"Scenario {scenario_name}_0 solved successfully.")
    elif params["create_matrix"]:
        print(f"Scenario {scenario_name}_0 matrix preparation completed successfully.")
    else:
        print(
            f"Scenario {scenario_name}_0 output stage SKIPPED; "
            "matrix and solver execution were disabled."
        )
    print(
        "\n#------------------------------------------------------------------"
        "------------#"
    )

    file_path_conv_format = os.path.join(
        here,
        params["Miscellaneous"],
        params["conv_format"],
    )
    file_path_template = os.path.join(
        here,
        params["A2_output_otoole"],
        scenario_name,
    )
    file_path_outputs = os.path.join(
        paths.folder_scenario,
        params["outputs"],
    )

    if solver == "glpk" and params["glpk_option"] == "new":
        output_command = [
            dependencies.get_executable("otoole"), "results", solver, "csv",
            f"{paths.output_file}.sol", file_path_outputs, "datafile",
            f"{paths.data_file}.txt", file_path_conv_format,
            "--glpk_model", f"{paths.output_file}.glp",
        ]
        if params["execute_model"]:
            dependencies.run_process(
                output_command,
                cwd=paths.folder_scenario,
                check=True,
            )
    elif solver in ["cbc", "cplex", "gurobi"]:
        output_command = [
            dependencies.get_executable("otoole"), "results", solver, "csv",
            f"{paths.output_file}.sol", file_path_outputs, "csv",
            file_path_template, file_path_conv_format,
        ]
        if params["execute_model"]:
            dependencies.run_process(
                output_command,
                cwd=paths.folder_scenario,
                check=True,
            )

    if solver in ["glpk", "cbc", "cplex", "gurobi"]:
        concatenate_command = [
            dependencies.python_executable,
            "-B",
            "-m",
            "ostram.pipeline.execution.concatenate",
            file_path_outputs, paths.output_file,
        ]
        if params["concat_otoole_csv"]:
            dependencies.run_process(
                concatenate_command,
                cwd=paths.folder_scenario,
                check=True,
            )
            print(
                f"Outputs concatenated to "
                f"{scenario_name}_0_Output.csv successfully."
            )
            print(
                "\n#---------------------------------------------------------------"
                "---------------#"
            )


def execute_scenario(
    params: dict[str, Any],
    scenario_name: str,
    here: Any,
    dependencies: ScenarioExecutionDependencies,
    *,
    solver_adapter: SolverAdapter | None = None,
    matrix_runner: Callable[[list[str], Callable[..., Any]], None] | None = None,
) -> None:
    """Plan and execute one scenario through explicit matrix/solver seams."""

    paths = resolve_scenario_execution_paths(params, scenario_name, here)
    solver = params["solver"]

    reuse_solution = (
        params.get("reuse_existing_sol", False)
        and dependencies.path_exists(paths.output_file + ".sol")
    )
    if params.get("reuse_existing_sol", False) and not reuse_solution:
        print(
            f"[reuse_existing_sol] Requested but {paths.output_file}.sol "
            "not found; falling back to a normal solve."
        )
    if reuse_solution:
        print(
            f"[reuse_existing_sol] Reusing existing solution: "
            f"{paths.output_file}.sol"
        )

    matrix_command: str | None = None
    solver_command: str | None = None

    if solver == "glpk":
        if params["execute_model"] and not reuse_solution:
            active_solver_adapter = solver_adapter or SolverAdapter(dependencies)
            solver_command = active_solver_adapter.prepare_command(
                solver,
                params,
                paths,
            )
    else:
        matrix_command = build_matrix_command(
            params,
            solver,
            paths,
            reuse_solution,
        )
        if solver in ["cbc", "cplex", "gurobi"]:
            if params["execute_model"] and not reuse_solution:
                active_solver_adapter = solver_adapter or SolverAdapter(dependencies)
                solver_command = active_solver_adapter.prepare_command(
                    solver,
                    params,
                    paths,
                )

    if params["execute_model"] or params["create_matrix"]:
        if matrix_command is not None:
            active_matrix_runner = matrix_runner or run_matrix_preparation
            active_matrix_runner(matrix_command, dependencies.run_process)
        if solver_command is not None:
            active_solver_adapter.invoke(solver_command)

    if (
        params["execute_model"]
        and solver in ["cbc", "cplex", "gurobi"]
        and not dependencies.path_exists(paths.output_file + ".sol")
    ):
        raise FileNotFoundError(
            "Solver finished but did not create the expected solution file: "
            f"{paths.output_file}.sol"
        )

    if params["execute_model"] and solver == "cbc":
        status = validate_cbc_solution(paths.output_file + ".sol")
        print(f"CBC solution status: {status}")

    run_scenario_output_stage(
        params,
        scenario_name,
        here,
        solver,
        paths,
        dependencies,
    )

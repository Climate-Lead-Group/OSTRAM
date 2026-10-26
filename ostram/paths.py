"""Central project, resource, and mutable-workspace resolution for OSTRAM.

Authoritative project content is resolved from an explicit project bundle;
caller current-working-directory is never used as an implicit resource root.
"""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import re
from typing import Mapping


PROJECT_ROOT_ENV = "OSTRAM_PROJECT_ROOT"
WORKSPACE_ENV = "OSTRAM_WORKSPACE"
SETUP_AUTHORITIES_ENV = "OSTRAM_SETUP_AUTHORITIES"
WINDOWS_SAFE_ABSOLUTE_PATH_BUDGET = 240
_WORKBOOK_PATH_DIGEST_LENGTH = 16
_SAFE_COMPONENT = re.compile(r"^[^./\\][^/\\]*$")
_SAFE_STAGE_IDENTITY = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]*$")


class ProjectResolutionError(RuntimeError):
    """Raised when no valid OSTRAM project bundle can be resolved."""


class WorkspacePathBudgetError(ValueError):
    """Raised when an ephemeral workbook cannot fit the safe path budget."""


def _absolute(path: str | os.PathLike[str]) -> Path:
    return Path(path).expanduser().resolve()


def windows_path_units(path: str | os.PathLike[str]) -> int:
    """Return the Windows path length in UTF-16 code units."""

    return len(os.fspath(path).encode("utf-16-le")) // 2


def bounded_workspace_workbook_path(
    desired_path: str | os.PathLike[str],
    *,
    stage_identity: str,
    budget: int = WINDOWS_SAFE_ABSOLUTE_PATH_BUDGET,
) -> Path:
    """Return a deterministic, Windows-safe path for a mutable workbook.

    Safe desired paths are returned unchanged after absolute resolution.  A
    path at or above ``budget`` is reduced to a meaningful final-stage label
    plus a digest of the complete desired absolute path.  The digest prevents
    two distinct over-budget desired names in the same workspace from being
    silently mapped to the same filename.
    """

    desired = _absolute(desired_path)
    if desired.suffix.lower() != ".xlsx":
        raise ValueError(
            "mutable workbook output must preserve the .xlsx extension: "
            f"{desired}"
        )
    if not _SAFE_STAGE_IDENTITY.fullmatch(stage_identity):
        raise ValueError(f"unsafe workbook stage identity: {stage_identity!r}")
    if budget <= 0:
        raise ValueError(f"path budget must be positive, got {budget}")
    if windows_path_units(desired) < budget:
        return desired

    normalized = os.path.normcase(os.fspath(desired))
    digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()[
        :_WORKBOOK_PATH_DIGEST_LENGTH
    ]
    compact_name = f"{stage_identity}_{digest}{desired.suffix}"
    compact = desired.with_name(compact_name)
    compact_units = windows_path_units(compact)
    if compact_units >= budget:
        parent_units = windows_path_units(desired.parent)
        available = budget - parent_units - 2
        raise WorkspacePathBudgetError(
            "mutable workspace parent leaves no Windows-safe workbook "
            f"filename budget: parent={desired.parent!s} "
            f"parent_length={parent_units} budget={budget} "
            f"available_filename_units={max(0, available)} "
            f"required_filename={compact_name!r}. The absolute output path "
            f"must be shorter than {budget} UTF-16 code units."
        )
    return compact


def _source_anchor() -> Path:
    return Path(__file__).resolve().parent.parent


def _layout_at(root: Path) -> str | None:
    valid = (
        (root / "ostram" / "__init__.py").is_file()
        and (root / "environment.yaml").is_file()
        and all((root / name).is_dir() for name in ("inputs", "config", "model"))
    )
    return "project" if valid else None


def _validated_root(candidate: str | os.PathLike[str], source: str) -> Path:
    root = _absolute(candidate)
    layout = _layout_at(root)
    if layout is None:
        raise ProjectResolutionError(
            f"{source} does not identify a valid OSTRAM project bundle: {root}. "
            "Expected the package and environment marker plus inputs/, "
            "config/, and model/."
        )
    return root


@dataclass(frozen=True)
class ProjectPaths:
    """Absolute paths for one validated project bundle and mutable workspace."""

    project_root: Path
    workspace: Path
    layout: str

    @classmethod
    def resolve(
        cls,
        *,
        project_root: str | os.PathLike[str] | None = None,
        workspace: str | os.PathLike[str] | None = None,
        environ: Mapping[str, str] | None = None,
    ) -> "ProjectPaths":
        """Resolve the project root and the workspace into one instance.

        The project root is ``project_root``, else ``OSTRAM_PROJECT_ROOT``, else
        the folder that holds the installed package. The workspace is
        ``workspace``, else ``OSTRAM_WORKSPACE``, else ``<project_root>/workspace``.

        Raises:
            ProjectResolutionError: If the chosen root is not a valid project
                bundle, or if no root is given and the package folder holds none.
        """

        environment = os.environ if environ is None else environ
        if project_root is not None:
            root = _validated_root(project_root, "--project-root")
        elif environment.get(PROJECT_ROOT_ENV):
            root = _validated_root(
                environment[PROJECT_ROOT_ENV], PROJECT_ROOT_ENV
            )
        else:
            anchor = _source_anchor()
            if _layout_at(anchor) is None:
                raise ProjectResolutionError(
                    "This non-editable OSTRAM installation has no colocated "
                    "project bundle. Supply --project-root or set "
                    f"{PROJECT_ROOT_ENV}."
                )
            root = anchor

        workspace_value = (
            workspace
            if workspace is not None
            else environment.get(WORKSPACE_ENV)
        )
        workspace_path = _absolute(
            workspace_value if workspace_value is not None else root / "workspace"
        )
        return cls(
            project_root=root,
            workspace=workspace_path,
            layout=_layout_at(root) or "",
        )

    @property
    def package_root(self) -> Path:
        """Return the package folder, ``<project_root>/ostram``."""

        return (self.project_root / "ostram").resolve()

    def authority(self, role: str, default: Path) -> Path:
        """Return an activated setup authority or the compatibility default.

        The canonical CLI injects the entire validated mapping in one
        environment value before importing a route.  Direct historical module
        calls have no mapping and retain their established default-setup paths.
        A present but incomplete mapping fails closed instead of falling back.
        """

        encoded = os.environ.get(SETUP_AUTHORITIES_ENV)
        if encoded is None:
            return default.resolve()
        try:
            mapping = json.loads(encoded)
        except json.JSONDecodeError as error:
            raise ProjectResolutionError(
                f"invalid {SETUP_AUTHORITIES_ENV}: {error}"
            ) from error
        if not isinstance(mapping, dict):
            raise ProjectResolutionError(
                f"{SETUP_AUTHORITIES_ENV} must encode an object"
            )
        if role not in mapping:
            raise ProjectResolutionError(
                f"active setup has no authority for role {role!r}"
            )
        value = mapping[role]
        if not isinstance(value, str) or not Path(value).is_absolute():
            raise ProjectResolutionError(
                f"active setup authority {role!r} is not an absolute path"
            )
        return Path(value).resolve()

    @property
    def inputs_root(self) -> Path:
        """Return the inputs folder, ``<project_root>/inputs``."""

        return (self.project_root / "inputs").resolve()

    @property
    def osemosys_inputs(self) -> Path:
        """Return the OSeMOSYS Global inputs folder.

        The default is ``<project_root>/inputs/osemosys_global``. The active
        setup's ``osemosys_inputs`` authority replaces it when a setup is
        active.
        """

        return self.authority(
            "osemosys_inputs", self.inputs_root / "osemosys_global"
        )

    @property
    def preparation_inputs(self) -> Path:
        """Return the A1 preparation inputs folder.

        The default is ``<project_root>/inputs/preparation``. The active
        setup's ``preparation_inputs`` authority replaces it when a setup is
        active.
        """

        return self.authority(
            "preparation_inputs", self.inputs_root / "preparation"
        )

    @property
    def preparation_templates(self) -> Path:
        """Return the workbook templates folder.

        The default is ``<preparation_inputs>/workbook_templates``. The active
        setup's ``preparation_templates`` authority replaces it when a setup
        is active.
        """

        return self.authority(
            "preparation_templates", self.preparation_inputs / "workbook_templates"
        )

    @property
    def secondary_technology_inputs(self) -> Path:
        """Return the secondary technology inputs folder.

        The default is
        ``<project_root>/inputs/preparation/secondary_technologies``. The active
        setup's ``secondary_technology_inputs`` authority replaces it when a
        setup is active.
        """

        return self.authority(
            "secondary_technology_inputs",
            self.inputs_root / "preparation" / "secondary_technologies",
        )

    @property
    def scenario_inputs(self) -> Path:
        """Return the scenario inputs folder.

        The default is ``<project_root>/inputs/scenarios``. The active setup's
        ``scenario_inputs`` authority replaces it when a setup is active.
        """

        return self.authority("scenario_inputs", self.inputs_root / "scenarios")

    @property
    def execution_inputs(self) -> Path:
        """Return the execution inputs folder.

        The default is ``<project_root>/inputs/execution``. The active setup's
        ``execution_inputs`` authority replaces it when a setup is active.
        """

        return self.authority("execution_inputs", self.inputs_root / "execution")

    @property
    def config_root(self) -> Path:
        """Return the configuration folder, ``<project_root>/config``."""

        return (self.project_root / "config").resolve()

    @property
    def scenario_config_root(self) -> Path:
        """Return the scenario configuration folder.

        The default is ``<project_root>/config/scenarios``. The active setup's
        ``scenario_config_root`` authority replaces it when a setup is active.
        """

        return self.authority(
            "scenario_config_root", self.config_root / "scenarios"
        )

    @property
    def preparation_config(self) -> Path:
        """Return the preparation configuration folder.

        The default is ``<project_root>/config/preparation``. The active
        setup's ``preparation_config`` authority replaces it when a setup is
        active.
        """

        return self.authority(
            "preparation_config", self.config_root / "preparation"
        )

    @property
    def country_config(self) -> Path:
        """Return the country code configuration.

        The default is
        ``<project_root>/config/preparation/Config_country_codes.yaml``. The
        active setup's ``country_config`` authority replaces it when a setup
        is active.
        """

        return self.authority(
            "country_config",
            self.config_root / "preparation" / "Config_country_codes.yaml",
        )

    @property
    def region_config(self) -> Path:
        """Return the region consolidation configuration.

        The default is
        ``<project_root>/config/preparation/Config_region_consolidation.yaml``.
        The active setup's ``region_config`` authority replaces it when a
        setup is active.
        """

        return self.authority(
            "region_config",
            self.config_root / "preparation" / "Config_region_consolidation.yaml",
        )

    @property
    def model_root(self) -> Path:
        """Return the model folder, ``<project_root>/model``."""

        return (self.project_root / "model").resolve()

    @property
    def maintained_model(self) -> Path:
        """Return the OSeMOSYS model file.

        The default is ``<project_root>/model/osemosys_fast_preprocessed.txt``.
        The active setup's ``maintained_model`` authority replaces it when a
        setup is active.
        """

        return self.authority(
            "maintained_model", self.model_root / "osemosys_fast_preprocessed.txt"
        )

    @property
    def package_resources_root(self) -> Path:
        """Return the package resources folder, ``<project_root>/ostram/resources``."""

        return (self.package_root / "resources").resolve()

    @property
    def compilation_resources(self) -> Path:
        """Return the compilation resources folder.

        The default is ``<project_root>/ostram/resources/compilation``. The
        active setup's ``compilation_resources`` authority replaces it when a
        setup is active.
        """

        return self.authority(
            "compilation_resources", self.package_resources_root / "compilation"
        )

    @property
    def preparation_resources(self) -> Path:
        """Return the preparation resources folder.

        The default is ``<project_root>/ostram/resources/preparation``. The
        active setup's ``preparation_resources`` authority replaces it when a
        setup is active.
        """

        return self.authority(
            "preparation_resources", self.package_resources_root / "preparation"
        )

    @property
    def scenario_registry(self) -> Path:
        """Return the scenario registry file.

        The default is the manifest of the default setup, as the setup
        registry names it. The active setup's ``scenario_registry`` authority
        replaces it when a setup is active.
        """

        from ostram.setups import DEFAULT_SETUP, load_setup

        return self.authority(
            "scenario_registry", load_setup(DEFAULT_SETUP, paths=self).path
        )

    @property
    def scenario_workbook(self) -> Path:
        """Return the scenario workbook.

        The default is
        ``<project_root>/inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx``. The
        active setup's ``scenario_workbook`` authority replaces it when a
        setup is active.
        """

        return self.authority(
            "scenario_workbook",
            self.inputs_root / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx",
        )

    @property
    def timeslice_workbook(self) -> Path:
        """Return the timeslice workbook.

        The default is
        ``<project_root>/inputs/scenarios/OSTRAM_Timeslice_Inputs.xlsx``. The
        active setup's ``timeslice_workbook`` authority replaces it when a
        setup is active.
        """

        return self.authority(
            "timeslice_workbook",
            self.inputs_root / "scenarios" / "OSTRAM_Timeslice_Inputs.xlsx",
        )

    @property
    def ao_decisions(self) -> Path:
        """Return the A-O decisions source.

        The default is
        ``<project_root>/inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx``. The
        active setup's ``ao_decisions`` authority replaces it when a setup
        is active.
        """

        return self.authority(
            "ao_decisions",
            self.inputs_root / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx",
        )

    @property
    def interconnector_authority(self) -> Path:
        """Return the interconnector source.

        The default is
        ``<project_root>/inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx``. The
        active setup's ``interconnector_authority`` authority replaces it when
        a setup is active.
        """

        return self.authority(
            "interconnector_authority",
            self.inputs_root / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx",
        )

    @property
    def interconnector_taxonomy(self) -> Path:
        """Return the technology taxonomy.

        The default is ``<project_root>/config/scenarios/technology_types.csv``.
        The active setup's ``interconnector_taxonomy`` authority replaces it
        when a setup is active.
        """

        return self.authority(
            "interconnector_taxonomy",
            self.config_root / "scenarios" / "technology_types.csv",
        )

    @property
    def execution_config(self) -> Path:
        """Return the execution configuration.

        The default is
        ``<project_root>/config/execution/Config_MOMF_T1_AB.yaml``. The active
        setup's ``execution_config`` authority replaces it when a setup is
        active.
        """

        return self.authority(
            "execution_config",
            self.config_root / "execution" / "Config_MOMF_T1_AB.yaml",
        )

    @property
    def compilation_config(self) -> Path:
        """Return the compilation configuration.

        The default is
        ``<project_root>/config/compilation/Config_MOMF_T1_A.yaml``. The active
        setup's ``compilation_config`` authority replaces it when a setup is
        active.
        """

        return self.authority(
            "compilation_config",
            self.config_root / "compilation" / "Config_MOMF_T1_A.yaml",
        )

    @property
    def preparation_workspace(self) -> Path:
        """Return the preparation workspace, ``<workspace>/preparation``."""

        return self.stage_workspace("preparation")

    @property
    def a1_outputs(self) -> Path:
        """Return the A1 output folder, ``<workspace>/preparation/A1_Outputs``."""

        return (self.preparation_workspace / "A1_Outputs").resolve()

    @property
    def generated_extra_inputs(self) -> Path:
        """Return the extra inputs folder, ``<workspace>/preparation/extra_inputs``."""

        return (self.preparation_workspace / "extra_inputs").resolve()

    @property
    def scenarios_workspace(self) -> Path:
        """Return the scenarios workspace, ``<workspace>/scenarios``."""

        return self.stage_workspace("scenarios")

    @property
    def compilation_workspace(self) -> Path:
        """Return the compilation workspace, ``<workspace>/compilation``."""

        return self.stage_workspace("compilation")

    @property
    def compiled_parameters(self) -> Path:
        """Return the B1 output folder, ``<workspace>/compilation/A2_Output_Params``."""

        return (self.compilation_workspace / "A2_Output_Params").resolve()

    @property
    def execution_workspace(self) -> Path:
        """Return the execution workspace, ``<workspace>/execution``."""

        return self.stage_workspace("execution")

    @property
    def otoole_outputs(self) -> Path:
        """Return the otoole output folder.

        The folder is ``<workspace>/execution/A2_Outputs_Params_otoole``.
        """

        return (self.execution_workspace / "A2_Outputs_Params_otoole").resolve()

    @property
    def executables(self) -> Path:
        """Return the executables folder, ``<workspace>/execution/Executables``."""

        return (self.execution_workspace / "Executables").resolve()

    @property
    def outputs(self) -> Path:
        """Return the results folder, ``<workspace>/execution/Outputs``."""

        return (self.execution_workspace / "Outputs").resolve()

    @property
    def environment_file(self) -> Path:
        """Return the conda environment file, ``<project_root>/environment.yaml``."""

        return (self.project_root / "environment.yaml").resolve()

    @property
    def dvc_file(self) -> Path:
        """Return the DVC pipeline file, ``<project_root>/dvc.yaml``."""

        return (self.project_root / "dvc.yaml").resolve()

    def resolve_project_file(
        self,
        value: str | os.PathLike[str],
        *,
        base: Path | None = None,
    ) -> Path:
        """Return ``value`` as an absolute path.

        A relative ``value`` is resolved against ``base``, or against the project
        root when ``base`` is ``None``.
        """

        path = Path(value).expanduser()
        if path.is_absolute():
            return path.resolve()
        return ((self.project_root if base is None else base) / path).resolve()

    def stage_workspace(
        self,
        stage: str,
        scenario: str | None = None,
        *,
        create: bool = False,
    ) -> Path:
        """Return the workspace folder of ``stage``, or of ``scenario`` within it.

        The folder is ``<workspace>/<stage>`` or
        ``<workspace>/<stage>/<scenario>``, and it is created when ``create`` is
        true.

        Raises:
            ValueError: If a component is empty, starts with a dot or contains a
                path separator.
        """

        components = [stage] if scenario is None else [stage, scenario]
        if any(not _SAFE_COMPONENT.fullmatch(component) for component in components):
            raise ValueError(f"unsafe workspace component: {components!r}")
        target = self.workspace.joinpath(*components).resolve()
        if create:
            target.mkdir(parents=True, exist_ok=True)
        return target

    def inspect_resources(self) -> dict[str, object]:
        """Read representative real resources without creating mutable state."""

        required = {
            "scenario_registry": self.scenario_registry,
            "scenario_workbook": self.scenario_workbook,
            "timeslice_workbook": self.timeslice_workbook,
            "compilation_config": self.compilation_config,
            "execution_config": self.execution_config,
            "maintained_model": self.maintained_model,
        }
        missing = [f"{name}={path}" for name, path in required.items() if not path.is_file()]
        if missing:
            raise FileNotFoundError("Required project resources are missing: " + ", ".join(missing))

        from ostram import yaml_io
        registry = yaml_io.load_file(self.scenario_registry)
        registry = registry.get("scenario_registry", registry)
        with self.scenario_workbook.open("rb") as stream:
            workbook_signature = stream.read(4).hex()
        config_first_line = next(
            line.strip()
            for line in self.execution_config.read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        )
        model_first_line = self.maintained_model.read_text(
            encoding="utf-8", errors="replace"
        ).splitlines()[0]
        return {
            "project_root": str(self.project_root),
            "workspace": str(self.workspace),
            "layout": self.layout,
            "registry_schema": registry.get("schema"),
            "root_scenarios": [entry["name"] for entry in registry["root_scenarios"]],
            "scenario_workbook": str(self.scenario_workbook),
            "scenario_workbook_signature": workbook_signature,
            "execution_config": str(self.execution_config),
            "execution_config_first_line": config_first_line,
            "maintained_model": str(self.maintained_model),
            "maintained_model_first_line": model_first_line,
            "package_resources": str(self.package_resources_root),
        }


def resolve_paths(
    *,
    project_root: str | os.PathLike[str] | None = None,
    workspace: str | os.PathLike[str] | None = None,
    environ: Mapping[str, str] | None = None,
) -> ProjectPaths:
    """Resolve the active project with CLI > environment > package precedence."""

    return ProjectPaths.resolve(
        project_root=project_root,
        workspace=workspace,
        environ=environ,
    )

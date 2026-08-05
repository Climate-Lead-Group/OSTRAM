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
PROFILE_AUTHORITIES_ENV = "OSTRAM_PROFILE_AUTHORITIES"
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
        return (self.project_root / "ostram").resolve()

    def authority(self, role: str, default: Path) -> Path:
        """Return an activated profile authority or the compatibility default.

        The canonical CLI injects the entire validated mapping in one
        environment value before importing a route.  Direct historical module
        calls have no mapping and retain their established full-model paths.
        A present but incomplete mapping fails closed instead of falling back.
        """

        encoded = os.environ.get(PROFILE_AUTHORITIES_ENV)
        if encoded is None:
            return default.resolve()
        try:
            mapping = json.loads(encoded)
        except json.JSONDecodeError as error:
            raise ProjectResolutionError(
                f"invalid {PROFILE_AUTHORITIES_ENV}: {error}"
            ) from error
        if not isinstance(mapping, dict):
            raise ProjectResolutionError(
                f"{PROFILE_AUTHORITIES_ENV} must encode an object"
            )
        if role not in mapping:
            raise ProjectResolutionError(
                f"active profile has no authority for role {role!r}"
            )
        value = mapping[role]
        if not isinstance(value, str) or not Path(value).is_absolute():
            raise ProjectResolutionError(
                f"active profile authority {role!r} is not an absolute path"
            )
        return Path(value).resolve()

    @property
    def inputs_root(self) -> Path:
        return (self.project_root / "inputs").resolve()

    @property
    def osemosys_inputs(self) -> Path:
        return self.authority(
            "osemosys_inputs", self.inputs_root / "osemosys_global"
        )

    @property
    def preparation_inputs(self) -> Path:
        return self.authority(
            "preparation_inputs", self.inputs_root / "preparation"
        )

    @property
    def preparation_templates(self) -> Path:
        return (self.preparation_inputs / "workbook_templates").resolve()

    @property
    def secondary_technology_inputs(self) -> Path:
        return self.authority(
            "secondary_technology_inputs",
            self.inputs_root / "preparation" / "secondary_technologies",
        )

    @property
    def scenario_inputs(self) -> Path:
        return self.authority("scenario_inputs", self.inputs_root / "scenarios")

    @property
    def execution_inputs(self) -> Path:
        return self.authority("execution_inputs", self.inputs_root / "execution")

    @property
    def config_root(self) -> Path:
        return (self.project_root / "config").resolve()

    @property
    def scenario_config_root(self) -> Path:
        return self.authority(
            "scenario_config_root", self.config_root / "scenarios"
        )

    @property
    def preparation_config(self) -> Path:
        return (self.config_root / "preparation").resolve()

    @property
    def country_config(self) -> Path:
        return self.authority(
            "country_config",
            self.config_root / "preparation" / "Config_country_codes.yaml",
        )

    @property
    def region_config(self) -> Path:
        return self.authority(
            "region_config",
            self.config_root / "preparation" / "Config_region_consolidation.yaml",
        )

    @property
    def model_root(self) -> Path:
        return (self.project_root / "model").resolve()

    @property
    def maintained_model(self) -> Path:
        return self.authority(
            "maintained_model", self.model_root / "osemosys_fast_preprocessed.txt"
        )

    @property
    def package_resources_root(self) -> Path:
        return (self.package_root / "resources").resolve()

    @property
    def compilation_resources(self) -> Path:
        return (self.package_resources_root / "compilation").resolve()

    @property
    def preparation_resources(self) -> Path:
        return (self.package_resources_root / "preparation").resolve()

    @property
    def scenario_registry(self) -> Path:
        return self.authority(
            "scenario_registry", self.config_root / "scenarios" / "registry.json"
        )

    @property
    def scenario_workbook(self) -> Path:
        return self.authority(
            "scenario_workbook",
            self.inputs_root / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx",
        )

    @property
    def timeslice_workbook(self) -> Path:
        return self.authority(
            "timeslice_workbook",
            self.inputs_root / "scenarios" / "OSTRAM_Timeslice_Inputs.xlsx",
        )

    @property
    def ao_decisions(self) -> Path:
        return self.authority(
            "ao_decisions",
            self.inputs_root / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx",
        )

    @property
    def interconnector_authority(self) -> Path:
        return self.authority(
            "interconnector_authority",
            self.inputs_root / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx",
        )

    @property
    def interconnector_taxonomy(self) -> Path:
        return self.authority(
            "interconnector_taxonomy",
            self.config_root / "scenarios" / "technology_types.csv",
        )

    @property
    def execution_config(self) -> Path:
        return self.authority(
            "execution_config",
            self.config_root / "execution" / "Config_MOMF_T1_AB.yaml",
        )

    @property
    def compilation_config(self) -> Path:
        return self.authority(
            "compilation_config",
            self.config_root / "compilation" / "Config_MOMF_T1_A.yaml",
        )

    @property
    def preparation_workspace(self) -> Path:
        return self.stage_workspace("preparation")

    @property
    def a1_outputs(self) -> Path:
        return (self.preparation_workspace / "A1_Outputs").resolve()

    @property
    def generated_extra_inputs(self) -> Path:
        return (self.preparation_workspace / "extra_inputs").resolve()

    @property
    def scenarios_workspace(self) -> Path:
        return self.stage_workspace("scenarios")

    @property
    def compilation_workspace(self) -> Path:
        return self.stage_workspace("compilation")

    @property
    def compiled_parameters(self) -> Path:
        return (self.compilation_workspace / "A2_Output_Params").resolve()

    @property
    def execution_workspace(self) -> Path:
        return self.stage_workspace("execution")

    @property
    def otoole_outputs(self) -> Path:
        return (self.execution_workspace / "A2_Outputs_Params_otoole").resolve()

    @property
    def executables(self) -> Path:
        return (self.execution_workspace / "Executables").resolve()

    @property
    def outputs(self) -> Path:
        return (self.execution_workspace / "Outputs").resolve()

    @property
    def environment_file(self) -> Path:
        return (self.project_root / "environment.yaml").resolve()

    @property
    def dvc_file(self) -> Path:
        return (self.project_root / "dvc.yaml").resolve()

    def resolve_project_file(
        self,
        value: str | os.PathLike[str],
        *,
        base: Path | None = None,
    ) -> Path:
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

        registry = json.loads(self.scenario_registry.read_text(encoding="utf-8"))
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

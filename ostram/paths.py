"""Central project, resource, and mutable-workspace resolution for OSTRAM.

Authoritative project content is resolved from an explicit project bundle;
caller current-working-directory is never used as an implicit resource root.
"""

from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
from typing import Mapping


PROJECT_ROOT_ENV = "OSTRAM_PROJECT_ROOT"
WORKSPACE_ENV = "OSTRAM_WORKSPACE"
_SAFE_COMPONENT = re.compile(r"^[^./\\][^/\\]*$")


class ProjectResolutionError(RuntimeError):
    """Raised when no valid OSTRAM project bundle can be resolved."""


def _absolute(path: str | os.PathLike[str]) -> Path:
    return Path(path).expanduser().resolve()


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

    @property
    def inputs_root(self) -> Path:
        return (self.project_root / "inputs").resolve()

    @property
    def osemosys_inputs(self) -> Path:
        return (self.inputs_root / "osemosys_global").resolve()

    @property
    def preparation_inputs(self) -> Path:
        return (self.inputs_root / "preparation").resolve()

    @property
    def preparation_templates(self) -> Path:
        return (self.preparation_inputs / "workbook_templates").resolve()

    @property
    def secondary_technology_inputs(self) -> Path:
        return (self.preparation_inputs / "secondary_technologies").resolve()

    @property
    def scenario_inputs(self) -> Path:
        return (self.inputs_root / "scenarios").resolve()

    @property
    def execution_inputs(self) -> Path:
        return (self.inputs_root / "execution").resolve()

    @property
    def config_root(self) -> Path:
        return (self.project_root / "config").resolve()

    @property
    def scenario_config_root(self) -> Path:
        return (self.config_root / "scenarios").resolve()

    @property
    def preparation_config(self) -> Path:
        return (self.config_root / "preparation").resolve()

    @property
    def model_root(self) -> Path:
        return (self.project_root / "model").resolve()

    @property
    def maintained_model(self) -> Path:
        return (self.model_root / "osemosys_fast_preprocessed.txt").resolve()

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
        return (self.config_root / "scenarios" / "registry.json").resolve()

    @property
    def scenario_workbook(self) -> Path:
        return (self.scenario_inputs / "OSTRAM_Scenario_Inputs.xlsx").resolve()

    @property
    def timeslice_workbook(self) -> Path:
        return (self.scenario_inputs / "OSTRAM_Timeslice_Inputs.xlsx").resolve()

    @property
    def execution_config(self) -> Path:
        return (self.config_root / "execution" / "Config_MOMF_T1_AB.yaml").resolve()

    @property
    def compilation_config(self) -> Path:
        return (self.config_root / "compilation" / "Config_MOMF_T1_A.yaml").resolve()

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

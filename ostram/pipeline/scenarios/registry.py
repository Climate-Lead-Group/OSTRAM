"""Canonical OSTRAM root and derived-scenario contract.

The registry is maintained in :mod:`scenario_registry.json`.  This module
validates that contract, resolves an exact canonical selection, and exposes the
root prerequisites needed by A1, A2, and A3.  It deliberately performs no
workbook or solver effects.
"""

from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
from typing import Iterable, Mapping, Sequence

from ostram.paths import resolve_paths

SUPPORT_SCENARIOS = ("BAU",)
ROOT_SCENARIOS = (
    "BAU",
    "A_Calibrated_BAU",
    "B_Optimised_VRE",
    "C_Target_VRE",
)
DECISION_SCENARIOS = (
    "A_Calibrated_BAU",
    "A_Calibrated_BAU_Clipped",
    "B_Optimised_VRE",
    "B_Opt_Clipped",
    "B_Opt_DirBidir",
    "B_Opt_DirContractual",
    "B_Opt_IndiaCosts",
    "B_Opt_IndiaCostsFuel",
    "B_Opt_SolarCapex130",
    "B_Opt_SolarCapexHi",
    "B_Opt_SolarCapexSpike",
    "B_Opt_TradeCap15",
    "B_Opt_TxCap150",
    "C_Target_VRE",
    "C_Target_VRE_Clipped",
)
CANONICAL_SCENARIOS = SUPPORT_SCENARIOS + DECISION_SCENARIOS
_SCENARIO_NAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]*$")
_ENVIRONMENT_NAME = re.compile(r"^[A-Z][A-Z0-9_]*$")


@dataclass(frozen=True)
class ResultDependency:
    """One completed-result dependency consumed during A3 materialization."""

    scenario: str
    environment: str
    default_path: Path


@dataclass(frozen=True)
class RootScenario:
    """One hand-maintained scenario from the workbook Control contract."""

    name: str
    role: str
    dependencies: tuple[ResultDependency, ...]


@dataclass(frozen=True)
class DerivedScenario:
    """One deterministic scenario derived from a maintained root."""

    name: str
    base_scenario: str
    patches: Path
    direction_overlay: Path | None = None
    direction_study_start_year: int | None = None


@dataclass(frozen=True)
class ScenarioRegistry:
    """Validated, immutable scenario registry."""

    path: Path
    support_scenarios: tuple[str, ...]
    roots: tuple[RootScenario, ...]
    derived: tuple[DerivedScenario, ...]
    decision_scenarios: tuple[str, ...]

    @property
    def root_names(self) -> tuple[str, ...]:
        return tuple(root.name for root in self.roots)

    @property
    def roots_by_name(self) -> dict[str, RootScenario]:
        return {root.name: root for root in self.roots}

    @property
    def derived_by_name(self) -> dict[str, DerivedScenario]:
        return {scenario.name: scenario for scenario in self.derived}

    @property
    def scenario_names(self) -> tuple[str, ...]:
        return self.support_scenarios + self.decision_scenarios

    def select(self, requested: str | Sequence[str] | None) -> tuple[str, ...]:
        """Return an exact selection in frozen canonical order.

        ``None`` selects BAU plus the accepted decision set.  An explicit
        comma-separated string or sequence is validated, de-duplicated
        fail-closed, and then ordered by the canonical contract.
        """

        if requested is None or requested == "":
            return self.scenario_names
        if isinstance(requested, str):
            names = [part.strip() for part in requested.split(",") if part.strip()]
        else:
            names = [str(part).strip() for part in requested if str(part).strip()]
        duplicates = sorted({name for name in names if names.count(name) > 1})
        if duplicates:
            raise ValueError(f"duplicate scenario selection: {duplicates}")
        unknown = [name for name in names if name not in self.scenario_names]
        if unknown:
            raise ValueError(
                f"unknown scenario selection: {unknown}. "
                f"Available: {list(self.scenario_names)}"
            )
        selected = set(names)
        return tuple(name for name in self.scenario_names if name in selected)

    def required_roots(self, selected: Iterable[str]) -> tuple[str, ...]:
        """Return root materialization prerequisites in Control order."""

        required: set[str] = set()
        derived = self.derived_by_name
        roots = set(self.root_names)
        for name in selected:
            if name in roots:
                required.add(name)
            elif name in derived:
                required.add(derived[name].base_scenario)
            else:
                raise ValueError(f"scenario is not registered: {name}")
        return tuple(name for name in self.root_names if name in required)

    def result_dependencies(
        self,
        roots: Iterable[str],
        *,
        execution_workspace: Path | None = None,
        environment: Mapping[str, str] | None = None,
    ) -> dict[str, Path]:
        """Resolve declared result dependencies without changing their source."""

        environ = os.environ if environment is None else environment
        workspace = (
            resolve_paths().execution_workspace
            if execution_workspace is None
            else execution_workspace
        )
        resolved: dict[str, Path] = {}
        roots_by_name = self.roots_by_name
        for root_name in roots:
            root = roots_by_name[root_name]
            for dependency in root.dependencies:
                override = environ.get(dependency.environment)
                path = (
                    Path(override).expanduser().resolve()
                    if override
                    else (workspace / dependency.default_path).resolve()
                )
                if not path.exists():
                    raise FileNotFoundError(
                        f"{root_name} requires a completed "
                        f"{dependency.scenario} result at {path}. Run the "
                        f"dependency first or set {dependency.environment}."
                    )
                resolved[dependency.environment] = path
        return resolved


def _duplicates(values: Sequence[str]) -> list[str]:
    return sorted({value for value in values if values.count(value) > 1})


def _registry_relative(base: Path, value: object, *, label: str) -> Path:
    raw = str(value)
    relative = Path(raw)
    if relative.is_absolute() or not raw or ".." in relative.parts:
        raise ValueError(f"unsafe {label} path: {raw!r}")
    resolved = (base / relative).resolve()
    try:
        resolved.relative_to(base.resolve())
    except ValueError as error:
        raise ValueError(f"{label} path escapes registry root: {raw!r}") from error
    return resolved


def load_registry(
    path: Path | str | None = None,
    *,
    validate_files: bool = True,
) -> ScenarioRegistry:
    """Load and fully validate the canonical scenario registry."""

    registry_path = (
        resolve_paths().scenario_registry
        if path is None
        else Path(path).resolve()
    )
    raw = json.loads(registry_path.read_text(encoding="utf-8"))
    if raw.get("schema") != "ostram-scenario-registry-v1":
        raise ValueError(f"unsupported scenario registry schema: {raw.get('schema')}")
    support_scenarios = tuple(str(name) for name in raw.get("support_scenarios", ()))

    base = registry_path.parent
    roots: list[RootScenario] = []
    for entry in raw.get("root_scenarios", ()):
        dependencies: list[ResultDependency] = []
        for dependency in entry.get("dependencies", ()):
            if dependency.get("type") != "result":
                raise ValueError(
                    f"unsupported dependency type for {entry.get('name')}: "
                    f"{dependency.get('type')}"
                )
            default_path = Path(str(dependency["default_path"]))
            if default_path.is_absolute() or ".." in default_path.parts:
                raise ValueError(
                    f"unsafe dependency default_path for {entry.get('name')}: "
                    f"{default_path}"
                )
            dependencies.append(
                ResultDependency(
                    scenario=str(dependency["scenario"]),
                    environment=str(dependency["environment"]),
                    default_path=default_path,
                )
            )
        roots.append(
            RootScenario(
                name=str(entry["name"]),
                role=str(entry["role"]),
                dependencies=tuple(dependencies),
            )
        )

    derived: list[DerivedScenario] = []
    for entry in raw.get("derived_scenarios", ()):
        overlay = entry.get("direction_overlay")
        patches = _registry_relative(
            base, entry["patches"], label=f"{entry.get('name')} patches"
        )
        direction_overlay = (
            _registry_relative(
                base, overlay, label=f"{entry.get('name')} direction overlay"
            )
            if overlay else None
        )
        derived.append(
            DerivedScenario(
                name=str(entry["name"]),
                base_scenario=str(entry["base_scenario"]),
                patches=patches,
                direction_overlay=direction_overlay,
                direction_study_start_year=entry.get(
                    "direction_study_start_year"
                ),
            )
        )

    decision_scenarios = tuple(raw.get("decision_scenarios", ()))
    registry = ScenarioRegistry(
        path=registry_path,
        support_scenarios=support_scenarios,
        roots=tuple(roots),
        derived=tuple(derived),
        decision_scenarios=decision_scenarios,
    )
    _validate_registry(registry, validate_files=validate_files)
    return registry


def _validate_registry(
    registry: ScenarioRegistry,
    *,
    validate_files: bool,
) -> None:
    root_names = list(registry.root_names)
    derived_names = [scenario.name for scenario in registry.derived]
    invalid_names = [
        name for name in root_names + derived_names
        if not _SCENARIO_NAME.fullmatch(name)
    ]
    if invalid_names:
        raise ValueError(f"unsafe registry scenario names: {invalid_names}")
    duplicates = _duplicates(root_names + derived_names)
    if duplicates:
        raise ValueError(f"duplicate registry scenario names: {duplicates}")

    if not registry.support_scenarios:
        raise ValueError("scenario registry must declare at least one support scenario")
    support_duplicates = _duplicates(list(registry.support_scenarios))
    if support_duplicates:
        raise ValueError(f"duplicate support scenarios: {support_duplicates}")
    if not set(registry.support_scenarios).issubset(root_names):
        raise ValueError("every support scenario must be a registered root")
    decision_duplicates = _duplicates(list(registry.decision_scenarios))
    if decision_duplicates:
        raise ValueError(f"duplicate decision scenarios: {decision_duplicates}")
    known = set(root_names) | set(derived_names)
    unknown_decisions = [
        name for name in registry.decision_scenarios if name not in known
    ]
    if unknown_decisions:
        raise ValueError(f"unknown decision scenarios: {unknown_decisions}")
    overlap = set(registry.support_scenarios) & set(registry.decision_scenarios)
    if overlap:
        raise ValueError(f"support and decision scenarios overlap: {sorted(overlap)}")

    expected_derived = [
        name for name in registry.decision_scenarios if name not in root_names
    ]
    if derived_names != expected_derived:
        raise ValueError(
            f"derived scenario order must be {expected_derived}, "
            f"got {derived_names}"
        )

    root_set = set(root_names)
    for scenario in registry.derived:
        if scenario.base_scenario not in root_set:
            raise ValueError(
                f"{scenario.name} has non-root base {scenario.base_scenario}"
            )
        if not validate_files:
            continue
        if not scenario.patches.is_file():
            raise FileNotFoundError(
                f"patch file missing for {scenario.name}: {scenario.patches}"
            )
        patch = json.loads(scenario.patches.read_text(encoding="utf-8"))
        if patch.get("scenario") != scenario.name:
            raise ValueError(
                f"{scenario.patches} declares scenario={patch.get('scenario')!r}; "
                f"expected {scenario.name!r}"
            )
        if patch.get("base_scenario") != scenario.base_scenario:
            raise ValueError(
                f"{scenario.patches} declares "
                f"base_scenario={patch.get('base_scenario')!r}; expected "
                f"{scenario.base_scenario!r}"
            )
        if (
            scenario.direction_overlay is not None
            and not scenario.direction_overlay.is_file()
        ):
            raise FileNotFoundError(
                f"direction overlay missing for {scenario.name}: "
                f"{scenario.direction_overlay}"
            )

    for root in registry.roots:
        for dependency in root.dependencies:
            if not _ENVIRONMENT_NAME.fullmatch(dependency.environment):
                raise ValueError(
                    f"unsafe dependency environment for {root.name}: "
                    f"{dependency.environment!r}"
                )
            if dependency.scenario not in root_set:
                raise ValueError(
                    f"{root.name} depends on unknown root "
                    f"{dependency.scenario}"
                )


def ensure_root_output_directories(
    a1_outputs: Path | str,
    registry: ScenarioRegistry | None = None,
) -> tuple[Path, ...]:
    """Create only the four canonical root output directories."""

    active_registry = load_registry() if registry is None else registry
    output_root = Path(a1_outputs)
    output_root.mkdir(parents=True, exist_ok=True)
    created: list[Path] = []
    for root in active_registry.root_names:
        path = output_root / f"A1_Outputs_{root}"
        path.mkdir(parents=True, exist_ok=True)
        created.append(path)
    return tuple(created)


def root_snapshots_exist(
    a1_outputs: Path | str,
    roots: Iterable[str],
) -> bool:
    """Return true only when every requested root has a post-A2 snapshot."""

    output_root = Path(a1_outputs)
    return all(
        (output_root / f"_post_a2_snapshot_{root}").is_dir()
        for root in roots
    )

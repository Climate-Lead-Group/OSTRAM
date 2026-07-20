"""Pure configuration and path planning for the legacy B1 compiler."""

from __future__ import annotations

from dataclasses import dataclass
import os
from typing import Any, MutableMapping


CONFIG_PATH = "Config_MOMF_T1_A.yaml"


@dataclass(frozen=True)
class TransformPlan:
    """Values derived before the legacy compiler starts its transformations.

    Paths deliberately remain strings and are resolved lazily.  The predecessor
    interprets every relative path against the process working directory and uses
    a mixture of ``os.path.join`` and string concatenation; normalizing those paths
    here would change that contract.
    """

    params: MutableMapping[str, Any]
    base_year: Any
    final_year: Any
    time_range_vector: list[int]
    wide_param_header: Any
    other_setup_params: dict[str, Any]
    other_setup_params_timeslices: list[Any]

    @property
    def years(self) -> list[int]:
        return self.time_range_vector

    @property
    def setup(self) -> dict[str, Any]:
        return self.other_setup_params

    def scenario_workbook(self, suffix_key: str) -> str:
        """Return the predecessor's exact nested A1 scenario path formula."""
        return os.path.join(
            self.params["A1_outputs"],
            self.params["A1_outputs"]
            + "_"
            + self.params["xtra_scen"]["Main_Scenario"]
            + self.params[suffix_key],
        )

    def extra_input(self, suffix_key: str) -> str:
        """Return an A2-extra-input path without normalizing its separators."""
        return self.params["A2_extra_inputs"] + self.params[suffix_key]

    @staticmethod
    def og_emissions_csv() -> str:
        """Return the fixed predecessor source used when ``Use_OG_module`` is on."""
        return os.path.join("OG_csvs_inputs", "EMISSION.csv")

    def structure_workbook(self) -> Any:
        return self.params["Print_A2_Struct_List"]

    def main_output_root(self) -> Any:
        return self.params["A2_output_main_scen"]

    def additional_output_root(self) -> Any:
        return self.params["A2_output"]


def build_transform_plan(params: MutableMapping[str, Any]) -> TransformPlan:
    """Derive the legacy plan while retaining its mutation and failure behavior.

    In particular, the configured ``Timeslices`` list is sorted in place.  The
    returned setup mapping is a new mapping whose values retain their predecessor
    identities, including that same sorted list.
    """
    base_year = params["base_year"]
    final_year = params["final_year"]
    time_range_vector = [
        year for year in range(int(base_year), int(final_year) + 1)
    ]
    wide_param_header = params["sets"]

    extra_scenarios = params["xtra_scen"]
    setup_names = list(extra_scenarios.keys())
    setup_values = list(extra_scenarios.values())
    configured_timeslices = setup_values[setup_names.index("Timeslices")]
    configured_timeslices.sort()

    other_setup_params: dict[str, Any] = {}
    for index in range(len(setup_names)):
        other_setup_params.update({setup_names[index]: setup_values[index]})

    return TransformPlan(
        params=params,
        base_year=base_year,
        final_year=final_year,
        time_range_vector=time_range_vector,
        wide_param_header=wide_param_header,
        other_setup_params=other_setup_params,
        other_setup_params_timeslices=configured_timeslices,
    )

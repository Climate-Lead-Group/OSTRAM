"""Pure configuration and path planning for the B1 compiler."""

from __future__ import annotations

from dataclasses import dataclass
import os
from typing import Any, MutableMapping

from ostram.paths import resolve_paths


CONFIG_PATH = "Config_MOMF_T1_A.yaml"


@dataclass(frozen=True)
class TransformPlan:
    """Values derived before the compiler starts its transformations.

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
    run_settings: dict[str, Any]  # the ``xtra_scen`` mapping
    run_settings_timeslices: list[Any]

    @property
    def years(self) -> list[int]:
        """Model years from the base year to the final year."""
        return self.time_range_vector

    def scenario_workbook(self, suffix_key: str) -> str:
        """Return the canonical nested A1 scenario workbook path."""
        output_root = self.params["A1_outputs"]
        scenario_prefix = os.path.basename(os.path.normpath(output_root))
        return os.path.join(
            output_root,
            scenario_prefix
            + "_"
            + self.params["xtra_scen"]["Main_Scenario"]
            + self.params[suffix_key],
        )

    def extra_input(self, suffix_key: str) -> str:
        """Return an A2-extra-input path without normalizing its separators."""
        return self.params["A2_extra_inputs"] + self.params[suffix_key]

    @staticmethod
    def og_emissions_csv() -> str:
        """Return the explicit mutable A1 copy used when ``Use_OG_module`` is on."""
        paths = resolve_paths()
        mutable = paths.preparation_workspace / "og_csvs_inputs" / "EMISSION.csv"
        return str(mutable if mutable.is_file() else paths.osemosys_inputs / "EMISSION.csv")

    def structure_workbook(self) -> Any:
        """Return the configured ``Print_A2_Struct_List`` workbook name."""
        return self.params["Print_A2_Struct_List"]

    def main_output_root(self) -> Any:
        """Return the configured ``A2_output_main_scen`` folder."""
        return self.params["A2_output_main_scen"]

    def additional_output_root(self) -> Any:
        """Return the configured ``A2_output`` folder."""
        return self.params["A2_output"]


def build_transform_plan(params: MutableMapping[str, Any]) -> TransformPlan:
    """Derive the plan while retaining the compiler's mutation and failure behavior.

    In particular, the configured ``Timeslices`` list is sorted in place.  The
    returned run settings mapping is a new mapping whose values retain their predecessor
    identities, including that same sorted list.
    """
    base_year = params["base_year"]
    final_year = params["final_year"]
    time_range_vector = [
        year for year in range(int(base_year), int(final_year) + 1)
    ]
    wide_param_header = params["sets"]

    extra_scenarios = params["xtra_scen"]
    run_setting_names = list(extra_scenarios.keys())
    run_setting_values = list(extra_scenarios.values())
    configured_timeslices = run_setting_values[run_setting_names.index("Timeslices")]
    configured_timeslices.sort()

    run_settings: dict[str, Any] = {}
    for index in range(len(run_setting_names)):
        run_settings.update({run_setting_names[index]: run_setting_values[index]})

    return TransformPlan(
        params=params,
        base_year=base_year,
        final_year=final_year,
        time_range_vector=time_range_vector,
        wide_param_header=wide_param_header,
        run_settings=run_settings,
        run_settings_timeslices=configured_timeslices,
    )

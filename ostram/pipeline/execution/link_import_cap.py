"""Serialise the import cap at existing link capacity during compilation.

The setup policy ``link_import_cap`` in ``config/setups/south_asia.yaml``
names the links, the exporting and importing countries and the first year.
The mode of each link is the one whose input fuel is in the exporting
country and whose output fuel is in the importing country, in the compiled
activity ratios of that year. The model file holds the constraint
``D4_LinkImportCap``; it reads only parameters the model already receives.
"""
from pathlib import Path

import pandas as pd

from ostram.setups import setup_policy

MODEL_SET = "set D4LinkImportCap"


def link_modes(input_folder, links, from_country, to_country, year):
    """Return {link: mode} for the direction from ``from_country`` to
    ``to_country``. Raises ValueError unless each link has exactly one such mode.
    """
    folder = Path(input_folder)
    ratios = {}
    for name in ("InputActivityRatio", "OutputActivityRatio"):
        df = pd.read_csv(folder / f"{name}.csv", dtype={"TECHNOLOGY": str, "FUEL": str})
        ratios[name] = df[(df["YEAR"].astype(int) == year) & (df["VALUE"].astype(float) > 0)]
    modes = {}
    for link in links:
        found = []
        for name, country in (("InputActivityRatio", from_country), ("OutputActivityRatio", to_country)):
            df = ratios[name]
            rows = df[(df["TECHNOLOGY"] == link) & df["FUEL"].str.startswith("ELC" + country)]
            found.append({int(m) for m in rows["MODE_OF_OPERATION"]})
        both = found[0] & found[1]
        if len(both) != 1:
            raise ValueError(
                f"Link import cap: {link} has {sorted(both)} modes from {from_country} to {to_country} in {year}"
            )
        modes[link] = both.pop()
    return modes


def render_link_cap(input_folder, model_text):
    """Return the GMPL set of the import cap at existing link capacity, or ''
    when the setup has no ``link_import_cap`` policy.

    Raises ValueError when the policy is set and the model file lacks the set.
    """
    policy = setup_policy("link_import_cap")
    if not policy:
        return ""
    if MODEL_SET not in model_text:
        raise ValueError(
            "The setup policy link_import_cap is set, but the model file lacks "
            f"'{MODEL_SET}', so the cap cannot be written to the solver datafile."
        )
    first = int(policy["first_year"])
    modes = link_modes(input_folder, policy["links"], policy["from_country"], policy["to_country"], first)
    flows = " ".join(f"({link},{mode},{first})" for link, mode in sorted(modes.items()))
    return f"{MODEL_SET} := {flows};\n"

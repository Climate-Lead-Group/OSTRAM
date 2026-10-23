"""Refuse rule configuration keys that no rule code reads."""

from __future__ import annotations

# Read by apply_internal_transmission and apply_internal_tx_losses from the
# shared ``internal_transmission`` block of Config_country_codes.yaml.
INTERNAL_TRANSMISSION_KEYS = frozenset({
    "esmap_capital_cost_usd2010_per_kw", "muv_2010", "muv_2022", "muv_2023",
    "fixed_cost_share", "gec2_cost_rs_crore", "gec2_capacity_mw", "inr_per_usd_2022",
    "operational_life", "transmission_loss", "residual_peak_margin",
})


def check_keys(mapping, allowed, path, section: str = "") -> None:
    """Raise ValueError naming ``path`` and every key of ``mapping`` outside ``allowed``."""
    unknown = sorted(str(key) for key in mapping if key not in allowed)
    if unknown:
        where = f"{path} ({section})" if section else str(path)
        raise ValueError(f"{where}: unknown keys: {', '.join(unknown)}")

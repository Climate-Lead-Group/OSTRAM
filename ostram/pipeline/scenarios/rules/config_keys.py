"""Refuse rule configuration keys that no rule code reads."""

from __future__ import annotations

# Read by apply_internal_transmission and apply_internal_tx_losses from the
# shared ``internal_transmission`` block of Config_country_codes.yaml.
INTERNAL_TRANSMISSION_KEYS = frozenset({
    "base_capital_cost", "base_fixed_cost", "re_capex_multiplier",
    "operational_life", "transmission_loss",
})


def check_keys(mapping, allowed, path, section: str = "") -> None:
    """Raise ValueError naming ``path`` and every key of ``mapping`` outside ``allowed``."""
    unknown = sorted(str(key) for key in mapping if key not in allowed)
    if unknown:
        where = f"{path} ({section})" if section else str(path)
        raise ValueError(f"{where}: unknown keys: {', '.join(unknown)}")

# -*- coding: utf-8 -*-
"""
Shared configuration loader for country and technology data.
Single source of truth: Config_country_codes.yaml
"""
import yaml
import unicodedata
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "Config_country_codes.yaml"

_cached_config = None


def _load_raw():
    global _cached_config
    if _cached_config is None:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            _cached_config = yaml.safe_load(f)
    return _cached_config


def strip_accents(text):
    """Remove accents from text for fuzzy country name matching (e.g., 'Haití' -> 'Haiti')"""
    nfkd = unicodedata.normalize('NFKD', str(text))
    return ''.join(c for c in nfkd if not unicodedata.combining(c))


# --- Country accessors ---

def get_country_data():
    """Returns raw country_data dict from YAML: {iso3: {english_name, ostram_name}}"""
    return _load_raw().get("country_data", {})


def get_countries():
    """Returns sorted list of active ISO-3 country codes."""
    return sorted(get_country_data().keys())


def get_country_names():
    """Returns {iso3: english_name} dict."""
    return {iso3: d["english_name"] for iso3, d in get_country_data().items()}


def get_iso_country_map():
    """Returns {iso3: english_name} + special entries (like INT). Used by A1, A2."""
    result = get_country_names()
    for code, name in _load_raw().get("special_entries", {}).items():
        result[code] = name
    return result


def get_ostram_country_mapping():
    """Returns {ostram_spanish_name: iso3} dict. Used for reading OSTRAM Excel files."""
    return {d["ostram_name"]: iso3 for iso3, d in get_country_data().items()}


def get_ostram_country_mapping_normalized():
    """Returns accent-stripped {name: iso3} dict for fuzzy matching."""
    return {strip_accents(name): iso3 for name, iso3 in get_ostram_country_mapping().items()}


def get_shares_country_mapping():
    """Returns {shares_name: iso3} dict. Used for Shares Excel files.
    Shares files use the same country names as OSTRAM files."""
    return get_ostram_country_mapping()


# --- Model settings ---

def get_first_year():
    """Returns the model reference/first year (e.g. 2023)."""
    return _load_raw().get("first_year", 2023)


def get_add_missing_countries_from_ostram():
    """Returns whether to add missing countries from OSTRAM generation data."""
    return _load_raw().get("add_missing_countries_from_ostram", False)


def get_pwr_cleanup_mode():
    """Returns PWR cleanup mode: 'drop', 'merge', or False (disabled)."""
    return _load_raw().get("pwr_cleanup_mode", False)


# --- Technology accessors ---

def get_ostram_tech_mapping():
    """Returns {ostram_tech_name: model_code} dict."""
    return _load_raw().get("ostram_tech_mapping", {})


def get_code_to_energy():
    """Returns {tech_code: description} dict."""
    return _load_raw().get("code_to_energy", {})


def get_renewable_fuels():
    """Returns set of renewable fuel codes."""
    return set(_load_raw().get("renewable_fuels", []))


def get_shares_tech_mapping():
    """Returns {shares_tech_name: model_code} dict."""
    return _load_raw().get("shares_tech_mapping", {})


def get_raw_config():
    """Returns the full raw YAML dict (for A2_AddTx transmission params)."""
    return _load_raw()

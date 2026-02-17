#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Aug 13 09:02:18 2025

@author: ClimateLeadGroup, Andrey Salazar-Vargas


Run scenarios by:
1) Discovering folders starting with 'A1_Outputs_' in the same directory as this script.
2) Building a list with the suffix after 'A1_Outputs_'.
3) Iterating the list:
   - Update xtra_scen.Main_Scenario in 'Config_MOMF_T1_A.yaml' to the current scenario.
   - Execute 'B1_Compiler.py'.
Notes:
- All referenced files are assumed to be in the same folder as this script.
- The YAML file is backed up before modifications and restored at the end.
"""

from pathlib import Path
import shutil
import subprocess
import sys
import re
from typing import List, Optional

# Try to use ruamel.yaml to preserve comments/formatting; fallback to PyYAML; final fallback to regex.
def try_import_yaml_handlers():
    ruamel_yaml = None
    pyyaml = None
    try:
        from ruamel.yaml import YAML  # type: ignore
        ruamel_yaml = YAML
    except Exception:
        ruamel_yaml = None
    if ruamel_yaml is None:
        try:
            import yaml  # type: ignore
            pyyaml = yaml
        except Exception:
            pyyaml = None
    return ruamel_yaml, pyyaml


def list_scenario_suffixes(base_dir: Path) -> List[str]:
    """Return list like ['BAU_NoRPO','NDC','NDC+ELC'] from folders 'A1_Outputs_*'."""
    suffixes: List[str] = []
    for item in sorted(base_dir.iterdir()):
        if item.is_dir() and item.name.startswith("A1_Outputs_"):
            suffix = item.name.split("A1_Outputs_", 1)[1]
            if suffix:  # Ensure non-empty
                suffixes.append(suffix)
    return suffixes


def read_yaml_ruamel(yaml_path: Path, YAML_cls):
    """Read YAML using ruamel.yaml to preserve formatting."""
    yaml = YAML_cls()
    yaml.preserve_quotes = True
    with yaml_path.open("r", encoding="utf-8") as f:
        data = yaml.load(f)
    return data, yaml


def write_yaml_ruamel(yaml_path: Path, data, yaml_obj):
    """Write YAML using ruamel.yaml."""
    with yaml_path.open("w", encoding="utf-8") as f:
        yaml_obj.dump(data, f)


def read_yaml_pyyaml(yaml_path: Path, pyyaml):
    """Read YAML using PyYAML (comments will be lost)."""
    with yaml_path.open("r", encoding="utf-8") as f:
        data = pyyaml.safe_load(f)
    return data


def write_yaml_pyyaml(yaml_path: Path, data, pyyaml):
    """Write YAML using PyYAML."""
    with yaml_path.open("w", encoding="utf-8") as f:
        pyyaml.safe_dump(data, f, sort_keys=False, allow_unicode=True)


def regex_update_main_scenario(yaml_text: str, new_value: str) -> str:
    """
    As a last resort, update the value of Main_Scenario within the xtra_scen block using regex.
    It tries to replace the first Main_Scenario line after 'xtra_scen:'.
    """
    # Locate xtra_scen block start
    xtra_match = re.search(r"(^|\n)xtra_scen:\s*\{?[\s\S]*?$", yaml_text)
    if not xtra_match:
        # Fallback: replace the first occurrence of Main_Scenario globally
        return re.sub(
            r"(Main_Scenario:\s*)['\"]?.*?['\"]?",
            rf"\1'{new_value}'",
            yaml_text,
            count=1,
        )

    # Replace Main_Scenario value (first occurrence after xtra_scen:)
    def replace_first_after_xtra(text: str) -> str:
        # We will just replace the first Main_Scenario occurrence anywhere;
        # still safe enough if the file contains it only under xtra_scen as expected.
        return re.sub(
            r"(Main_Scenario:\s*)['\"]?.*?['\"]?",
            rf"\1'{new_value}'",
            text,
            count=1,
        )

    return replace_first_after_xtra(yaml_text)


def update_main_scenario(yaml_path: Path, new_value: str) -> None:
    """Update xtra_scen.Main_Scenario in the YAML file using the best available method."""
    ruamel_yaml_cls, pyyaml_mod = try_import_yaml_handlers()

    if ruamel_yaml_cls is not None:
        # Preferred path: preserve comments/formatting
        data, yaml_obj = read_yaml_ruamel(yaml_path, ruamel_yaml_cls)
        if not isinstance(data, dict) or "xtra_scen" not in data:
            raise ValueError("YAML does not contain 'xtra_scen' at the top level.")
        if not isinstance(data["xtra_scen"], dict):
            raise ValueError("'xtra_scen' is not a mapping in the YAML.")
        data["xtra_scen"]["Main_Scenario"] = new_value
        write_yaml_ruamel(yaml_path, data, yaml_obj)
        return

    if pyyaml_mod is not None:
        # Fallback: PyYAML (comments will be lost)
        data = read_yaml_pyyaml(yaml_path, pyyaml_mod)
        if not isinstance(data, dict) or "xtra_scen" not in data or not isinstance(data["xtra_scen"], dict):
            raise ValueError("YAML structure invalid or missing 'xtra_scen'.")
        data["xtra_scen"]["Main_Scenario"] = new_value
        write_yaml_pyyaml(yaml_path, data, pyyaml_mod)
        return

    # Last resort: regex line replacement
    original_text = yaml_path.read_text(encoding="utf-8")
    updated_text = regex_update_main_scenario(original_text, new_value)
    yaml_path.write_text(updated_text, encoding="utf-8")


def run_compiler(script_dir: Path) -> int:
    """Execute B1_Compiler.py with the current Python interpreter."""
    compiler = script_dir / "B1_Compiler.py"
    if not compiler.is_file():
        raise FileNotFoundError(f"Missing script: {compiler}")
    # Run using same Python interpreter
    result = subprocess.run([sys.executable, str(compiler)], cwd=str(script_dir))
    return result.returncode


def main():
    # Resolve base directory (same folder as this script)
    script_dir = Path(__file__).resolve().parent

    # Define key paths
    yaml_file = script_dir / "Config_MOMF_T1_A.yaml"
    compiler_script = script_dir / "B1_Compiler.py"
    A1_Outputs_script = script_dir / "A1_Outputs"


    if not yaml_file.is_file():
        print(f"[ERROR] YAML not found: {yaml_file}")
        sys.exit(1)
    if not compiler_script.is_file():
        print(f"[ERROR] Compiler script not found: {compiler_script}")
        sys.exit(1)

    # Build scenario list from folder names
    scenario_suffixes = list_scenario_suffixes(A1_Outputs_script)
    if not scenario_suffixes:
        print("[WARN] No 'A1_Outputs_*' folders found. Nothing to do.")
        sys.exit(0)

    print(f"[INFO] Scenarios discovered: {scenario_suffixes}")

    # Backup YAML
    backup_path = yaml_file.with_suffix(yaml_file.suffix + ".bak")
    shutil.copy2(yaml_file, backup_path)
    print(f"[INFO] Backup created: {backup_path.name}")

    # Iterate scenarios
    try:
        for scenario in scenario_suffixes:
            print(f"\n[INFO] === Running scenario: {scenario} ===")
            # 1) Update YAML
            try:
                update_main_scenario(yaml_file, scenario)
                print(f"[INFO] Updated 'Main_Scenario' to '{scenario}' in {yaml_file.name}")
            except Exception as e:
                print(f"[ERROR] Failed to update YAML for scenario '{scenario}': {e}")
                # Continue to next scenario but keep trying others
                continue

            # 2) Execute compiler
            rc = run_compiler(script_dir)
            if rc != 0:
                print(f"[ERROR] B1_Compiler.py exited with code {rc} for scenario '{scenario}'")
            else:
                print(f"[INFO] B1_Compiler.py completed successfully for scenario '{scenario}'")

    finally:
        # Restore original YAML from backup
        try:
            shutil.move(str(backup_path), str(yaml_file))
            print(f"\n[INFO] Restored original YAML from backup.")
        except Exception as e:
            print(f"\n[WARN] Could not restore YAML from backup: {e}")
            # If move failed, try to at least keep a copy
            if backup_path.exists():
                print(f"[WARN] Backup still available at: {backup_path}")

    print("\n[INFO] All done.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Public CLI for behavior-preserving B1 compiler orchestration."""

from __future__ import annotations


try:
    from t1_confection import b1_runner as _impl
except ModuleNotFoundError as error:
    if error.name != "t1_confection":
        raise
    import b1_runner as _impl


# Preserve the existing import surface while the implementation lives behind the CLI.
try_import_yaml_handlers = _impl.try_import_yaml_handlers
list_scenario_suffixes = _impl.list_scenario_suffixes
read_yaml_ruamel = _impl.read_yaml_ruamel
write_yaml_ruamel = _impl.write_yaml_ruamel
read_yaml_pyyaml = _impl.read_yaml_pyyaml
write_yaml_pyyaml = _impl.write_yaml_pyyaml
regex_update_main_scenario = _impl.regex_update_main_scenario
update_main_scenario = _impl.update_main_scenario
run_compiler = _impl.run_compiler
parse_cli_args = _impl.parse_cli_args


def main() -> None:
    return _impl.orchestrate(
        parse_cli_args(),
        _impl.B1Paths.from_entrypoint(__file__),
        scenario_discoverer=list_scenario_suffixes,
        scenario_updater=update_main_scenario,
        compiler_runner=run_compiler,
    )


if __name__ == "__main__":
    main()

"""Installable, caller-CWD-independent command interface for OSTRAM."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
from dataclasses import dataclass
import importlib
import json
import os
from pathlib import Path
import subprocess
import sys
from types import ModuleType
from typing import Iterator, Sequence

from ostram.paths import PROJECT_ROOT_ENV, WORKSPACE_ENV, ProjectPaths, resolve_paths
from ostram.setup_workspace import prepared_setup, setup_workspace
from ostram.setups import (
    DEFAULT_SETUP,
    SetupManifest,
    encoded_environment,
    load_setup,
)
from ostram.terminal import safe_print


@dataclass(frozen=True)
class Route:
    """One canonical command mapped to its package implementation."""

    module_name: str
    program: str
    help: str
    exit_policy: str


ROUTES = {
    "run": Route(
        module_name="ostram.pipeline.orchestration",
        program="python -m ostram run",
        help=(
            "Prepare the base model (A1), add the transmission network (A2), "
            "build scenarios (A3), compile inputs (B1), and run/collect "
            "results (B2)."
        ),
        exit_policy="run-guard",
    ),
    "transform": Route(
        module_name="ostram.pipeline.scenarios.transform",
        program="python -m ostram transform",
        help="Run the established A3 transformation for one scenario.",
        exit_policy="main-result",
    ),
    "compile-inputs": Route(
        module_name="ostram.pipeline.compilation.runner",
        program="python -m ostram compile-inputs",
        help="Run the established B1 multi-scenario compiler runner.",
        exit_policy="natural-zero",
    ),
}

EXTENDED_ROUTES = {
    "example": Route(
        module_name="ostram.examples",
        program="python -m ostram example",
        help="Prepare or report a registered example setup.",
        exit_policy="main-result",
    ),
    "country": Route(
        module_name="ostram.pipeline.preparation.country_commands",
        program="python -m ostram country",
        help="Generate, merge, or validate country data for the active setup.",
        exit_policy="main-result",
    ),
    "scenario": Route(
        module_name="ostram.pipeline.preparation.scenario_country_sync",
        program="python -m ostram scenario",
        help="Run setup-aware scenario preparation operations.",
        exit_policy="main-result",
    ),
    "new-scenario": Route(
        module_name="ostram.pipeline.scenarios.new_scenario",
        program="python -m ostram new-scenario",
        help="Create a derived scenario as a copy of an existing one.",
        exit_policy="main-result",
    ),
}


def build_parser() -> argparse.ArgumentParser:
    """Build the top-level argument parser for ``python -m ostram``."""
    parser = argparse.ArgumentParser(
        prog="python -m ostram",
        description="Canonical command interface for the OSTRAM workflow.",
        epilog=(
            "Project resources never come from caller CWD. Place global "
            "--project-root/--workspace options before COMMAND."
        ),
    )
    parser.add_argument(
        "--project-root",
        help=(
            "OSTRAM project bundle (precedence over OSTRAM_PROJECT_ROOT and "
            "the validated editable-checkout anchor)."
        ),
    )
    parser.add_argument(
        "--workspace",
        help=(
            "Mutable workspace (precedence over OSTRAM_WORKSPACE and "
            "<project-root>/workspace)."
        ),
    )
    parser.add_argument(
        "--setup",
        default=DEFAULT_SETUP,
        help="Fail-closed OSTRAM setup (default: south_asia).",
    )
    subcommands = parser.add_subparsers(dest="command", metavar="COMMAND")
    for name, route in ROUTES.items():
        subcommands.add_parser(name, add_help=False, help=route.help)
    for name, route in EXTENDED_ROUTES.items():
        subcommands.add_parser(name, add_help=False, help=route.help)
    subcommands.add_parser(
        "inspect-resources",
        add_help=False,
        help="Read and report real project input/config/model resources safely.",
    )
    return parser


def _global_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--project-root")
    parser.add_argument("--workspace")
    parser.add_argument("--setup", default=DEFAULT_SETUP)
    return parser


@contextmanager
def _historical_argv(program: str, arguments: Sequence[str]) -> Iterator[None]:
    previous = sys.argv
    sys.argv = [program, *arguments]
    try:
        yield
    finally:
        sys.argv = previous


@contextmanager
def _resolved_environment(
    paths: ProjectPaths,
    setup_environment: dict[str, str] | None = None,
    *,
    effective_workspace: Path | None = None,
) -> Iterator[None]:
    setup_environment = {} if setup_environment is None else setup_environment
    previous = {
        PROJECT_ROOT_ENV: os.environ.get(PROJECT_ROOT_ENV),
        WORKSPACE_ENV: os.environ.get(WORKSPACE_ENV),
        **{name: os.environ.get(name) for name in setup_environment},
    }
    os.environ[PROJECT_ROOT_ENV] = str(paths.project_root)
    os.environ[WORKSPACE_ENV] = str(
        paths.workspace if effective_workspace is None else effective_workspace
    )
    os.environ.update(setup_environment)
    try:
        yield
    finally:
        for name, value in previous.items():
            if value is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = value


def _load_route_module(route: Route) -> ModuleType:
    return importlib.import_module(route.module_name)


def _invoke_run_guard(module: ModuleType) -> int:
    try:
        result = module.main()
    except KeyboardInterrupt:
        safe_print("\nOSTRAM run interrupted", file=sys.stderr)
        return 130
    except subprocess.CalledProcessError as error:
        safe_print(
            f"\nCommand failed (exit {error.returncode}): {error.cmd}",
            file=sys.stderr,
        )
        return error.returncode
    except Exception as error:
        safe_print(f"\nError: {error}", file=sys.stderr)
        return 1
    return result if isinstance(result, int) else 0


def _invoke_route(route: Route, arguments: Sequence[str]) -> int | None:
    with _historical_argv(route.program, arguments):
        module = _load_route_module(route)
        if route.exit_policy == "run-guard":
            return _invoke_run_guard(module)
        result = module.main()
        if route.exit_policy == "main-result":
            return result
        return 0


def _resolve_cli(
    arguments: list[str],
) -> tuple[ProjectPaths, argparse.Namespace, list[str]]:
    """Parse only the global prefix so options never leak to child parsers."""

    known = {"--project-root", "--workspace", "--setup"}
    prefix: list[str] = []
    index = 0
    while index < len(arguments):
        token = arguments[index]
        option = token.split("=", 1)[0]
        if option not in known:
            break
        prefix.append(token)
        if "=" not in token:
            index += 1
            if index >= len(arguments):
                # Let argparse produce its established missing-value error.
                break
            prefix.append(arguments[index])
        index += 1
    options = _global_parser().parse_args(prefix)
    options.setup_explicit = any(
        token == "--setup" or token.startswith("--setup=") for token in prefix
    )
    remaining = arguments[index:]
    return (
        resolve_paths(
            project_root=options.project_root,
            workspace=options.workspace,
        ),
        options,
        remaining,
    )


def _selection_id(options: argparse.Namespace, remaining: Sequence[str]) -> str:
    selected = options.setup
    if len(remaining) >= 3 and remaining[0] == "example" and remaining[1] in {
        "prepare",
        "report",
    }:
        positional = remaining[2]
        if options.setup_explicit and selected != positional:
            raise ValueError(
                f"--setup {selected!r} conflicts with example setup {positional!r}"
            )
        selected = positional
    return selected


def _activate_setup(
    setup_id: str,
    *,
    paths: ProjectPaths,
    preparing: bool,
) -> tuple[SetupManifest, dict[str, str], Path]:
    manifest = load_setup(setup_id, paths=paths)
    if preparing:
        authorities = manifest.source_paths(paths, require_exists=True)
        prepared_workspace = setup_workspace(paths, setup_id)
        active_workspace = paths.workspace
    else:
        prepared = prepared_setup(manifest, paths=paths)
        authorities = prepared.authorities
        active_workspace = (
            paths.workspace if setup_id == DEFAULT_SETUP else prepared.workspace
        )
        prepared_workspace = active_workspace
    environment = encoded_environment(
        manifest,
        authorities=authorities,
        setup_workspace=prepared_workspace,
    )
    return manifest, environment, active_workspace


def main(argv: Sequence[str] | None = None) -> int | None:
    """Resolve the project paths and setup, then dispatch the command to its route."""
    arguments = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    if not arguments:
        parser.print_help()
        return 0
    if arguments == ["--help"]:
        parser.print_help()
        return 0

    try:
        paths, options, remaining = _resolve_cli(arguments)
    except SystemExit:
        raise
    except Exception as error:
        parser.error(str(error))

    if not remaining:
        parser.error("a command is required")
    command = remaining[0]
    try:
        setup_id = _selection_id(options, remaining)
        preparing = (
            len(remaining) >= 2
            and command == "example"
            and remaining[1] == "prepare"
        )
        # Help is a source-inspection operation. It must remain available in a
        # fresh checkout before a mutable setup workspace has been prepared.
        preparing = preparing or any(
            token in {"-h", "--help"} for token in remaining[1:]
        )
        _, setup_environment, active_workspace = _activate_setup(
            setup_id,
            paths=paths,
            preparing=preparing,
        )
    except Exception as error:
        parser.error(str(error))

    with _resolved_environment(
        paths,
        setup_environment,
        effective_workspace=active_workspace,
    ):
        if command in ROUTES:
            return _invoke_route(ROUTES[command], remaining[1:])
        if command in EXTENDED_ROUTES:
            return _invoke_route(EXTENDED_ROUTES[command], remaining[1:])
        if command == "inspect-resources":
            if len(remaining) != 1:
                parser.error("inspect-resources accepts no command arguments")
            # Re-resolve after setup activation so non-default setups report
            # their isolated prepared workspace rather than the outer workspace
            # used only to locate ``setups/<id>``.
            print(
                json.dumps(
                    resolve_paths().inspect_resources(), indent=2, sort_keys=True
                )
            )
            return 0

    parser.parse_args(arguments)
    parser.error("a command is required")


if __name__ == "__main__":
    raise SystemExit(main())

"""Installable, caller-CWD-independent command interface for OSTRAM."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
from dataclasses import dataclass
import importlib
import json
import os
import subprocess
import sys
from types import ModuleType
from typing import Iterator, Sequence

from ostram.paths import PROJECT_ROOT_ENV, WORKSPACE_ENV, ProjectPaths, resolve_paths


@dataclass(frozen=True)
class Route:
    """One canonical name mapped to an unchanged Stage 11 entrypoint."""

    module_name: str
    program: str
    help: str
    exit_policy: str


ROUTES = {
    "run": Route(
        module_name="run",
        program="run.py",
        help="Run the established A1/A2/A3/B1/B2 orchestration.",
        exit_policy="run-guard",
    ),
    "transform": Route(
        module_name="t1_confection.A3_process",
        program="A3_process.py",
        help="Run the established A3 transformation for one scenario.",
        exit_policy="main-result",
    ),
    "compile-inputs": Route(
        module_name="t1_confection.B1_Run_Compiler",
        program="B1_Run_Compiler.py",
        help="Run the established B1 multi-scenario compiler runner.",
        exit_policy="natural-zero",
    ),
}


def build_parser() -> argparse.ArgumentParser:
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
    subcommands = parser.add_subparsers(dest="command", metavar="COMMAND")
    for name, route in ROUTES.items():
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
def _resolved_environment(paths: ProjectPaths) -> Iterator[None]:
    previous = {
        PROJECT_ROOT_ENV: os.environ.get(PROJECT_ROOT_ENV),
        WORKSPACE_ENV: os.environ.get(WORKSPACE_ENV),
    }
    os.environ[PROJECT_ROOT_ENV] = str(paths.project_root)
    os.environ[WORKSPACE_ENV] = str(paths.workspace)
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
        module.main()
    except subprocess.CalledProcessError as error:
        print(
            f"\nCommand failed (exit {error.returncode}): {error.cmd}",
            file=sys.stderr,
        )
        return error.returncode
    except Exception as error:
        print(f"\nError: {error}", file=sys.stderr)
        return 1
    return 0


def _invoke_route(route: Route, arguments: Sequence[str]) -> int | None:
    with _historical_argv(route.program, arguments):
        module = _load_route_module(route)
        if route.exit_policy == "run-guard":
            return _invoke_run_guard(module)
        result = module.main()
        if route.exit_policy == "main-result":
            return result
        return 0


def _resolve_cli(arguments: list[str]) -> tuple[ProjectPaths, list[str]]:
    options, remaining = _global_parser().parse_known_args(arguments)
    return (
        resolve_paths(
            project_root=options.project_root,
            workspace=options.workspace,
        ),
        remaining,
    )


def main(argv: Sequence[str] | None = None) -> int | None:
    arguments = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    if not arguments:
        parser.print_help()
        return 0
    if arguments == ["--help"]:
        parser.print_help()
        return 0

    try:
        paths, remaining = _resolve_cli(arguments)
    except SystemExit:
        raise
    except Exception as error:
        parser.error(str(error))

    if not remaining:
        parser.error("a command is required")
    command = remaining[0]
    with _resolved_environment(paths):
        if command in ROUTES:
            return _invoke_route(ROUTES[command], remaining[1:])
        if command == "inspect-resources":
            if len(remaining) != 1:
                parser.error("inspect-resources accepts no command arguments")
            print(json.dumps(paths.inspect_resources(), indent=2, sort_keys=True))
            return 0

    parser.parse_args(arguments)
    parser.error("a command is required")


if __name__ == "__main__":
    raise SystemExit(main())

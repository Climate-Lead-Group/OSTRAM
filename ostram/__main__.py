"""Repository-local, platform-neutral command interface for OSTRAM."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
from dataclasses import dataclass
import importlib
import subprocess
import sys
from types import ModuleType
from typing import Iterator, Sequence


@dataclass(frozen=True)
class Route:
    """One canonical name mapped to an unchanged historical entrypoint."""

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
    """Build only the canonical hierarchy; route parsers remain historical."""

    parser = argparse.ArgumentParser(
        prog="python -m ostram",
        description="Canonical command interface for the OSTRAM workflow.",
        epilog=(
            "Historical script commands remain supported compatibility "
            "entrypoints. Run a subcommand with --help for its established "
            "arguments."
        ),
    )
    subcommands = parser.add_subparsers(dest="command", metavar="COMMAND")
    for name, route in ROUTES.items():
        subcommands.add_parser(name, add_help=False, help=route.help)
    return parser


@contextmanager
def _historical_argv(program: str, arguments: Sequence[str]) -> Iterator[None]:
    """Expose exact downstream argv temporarily and restore the original object."""

    previous = sys.argv
    sys.argv = [program, *arguments]
    try:
        yield
    finally:
        sys.argv = previous


def _load_route_module(route: Route) -> ModuleType:
    """Lazy-load only the selected import-safe historical entrypoint."""

    return importlib.import_module(route.module_name)


def _invoke_run_guard(module: ModuleType) -> int:
    """Preserve ``run.py``'s direct-script exception and exit translation."""

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
        # Import failures happen before run.py's historical guard and must not be
        # translated as launcher failures.
        module = _load_route_module(route)
        if route.exit_policy == "run-guard":
            return _invoke_run_guard(module)

        result = module.main()
        if route.exit_policy == "main-result":
            return result
        return 0


def main(argv: Sequence[str] | None = None) -> int | None:
    """Dispatch one canonical command without parsing its downstream arguments."""

    arguments = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    if not arguments:
        parser.print_help()
        return 0

    command = arguments[0]
    if command in ROUTES:
        return _invoke_route(ROUTES[command], arguments[1:])

    # Let argparse retain standard stdout/stderr and exit 0/2 behavior for the
    # canonical help option and malformed or unknown top-level commands.
    parser.parse_args(arguments)
    parser.error("a command is required")


if __name__ == "__main__":
    raise SystemExit(main())

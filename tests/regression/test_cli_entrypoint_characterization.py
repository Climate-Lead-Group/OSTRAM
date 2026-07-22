from __future__ import annotations

import ast
import importlib
import io
import sys
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]

RUN = REPO_ROOT / "run.py"
A3 = REPO_ROOT / "t1_confection" / "A3_process.py"
B1_RUNNER = REPO_ROOT / "t1_confection" / "B1_Run_Compiler.py"
B1_COMPILER = REPO_ROOT / "t1_confection" / "B1_Compiler.py"
B2 = REPO_ROOT / "t1_confection" / "B2_Executing_OG_Model.py"
B1_IMPLEMENTATION = REPO_ROOT / "t1_confection" / "b1_runner.py"
B2_IMPLEMENTATION = REPO_ROOT / "t1_confection" / "b2_orchestrator.py"

HISTORICAL_ENTRYPOINTS = (RUN, A3, B1_RUNNER, B1_COMPILER, B2)


def _source(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def _tree(path: Path) -> ast.Module:
    return ast.parse(_source(path), filename=str(path))


def _functions(path: Path) -> dict[str, ast.FunctionDef | ast.AsyncFunctionDef]:
    return {
        node.name: node
        for node in _tree(path).body
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }


def _main_guard(path: Path) -> ast.If | None:
    for node in _tree(path).body:
        if not isinstance(node, ast.If):
            continue
        test = node.test
        if (
            isinstance(test, ast.Compare)
            and isinstance(test.left, ast.Name)
            and test.left.id == "__name__"
            and len(test.ops) == 1
            and isinstance(test.ops[0], ast.Eq)
            and len(test.comparators) == 1
            and isinstance(test.comparators[0], ast.Constant)
            and test.comparators[0].value == "__main__"
        ):
            return node
    return None


def _capture_exit(call):
    stdout = io.StringIO()
    stderr = io.StringIO()
    with redirect_stdout(stdout), redirect_stderr(stderr):
        try:
            result = call()
        except SystemExit as error:
            code = error.code
        else:
            code = result
    return code, stdout.getvalue(), stderr.getvalue()


class HistoricalEntrypointInventoryTests(unittest.TestCase):
    def test_exact_historical_paths_exist_and_parse(self) -> None:
        self.assertEqual(
            [path.relative_to(REPO_ROOT).as_posix() for path in HISTORICAL_ENTRYPOINTS],
            [
                "run.py",
                "t1_confection/A3_process.py",
                "t1_confection/B1_Run_Compiler.py",
                "t1_confection/B1_Compiler.py",
                "t1_confection/B2_Executing_OG_Model.py",
            ],
        )
        for path in HISTORICAL_ENTRYPOINTS:
            with self.subTest(path=path):
                self.assertTrue(path.is_file())
                self.assertIsInstance(_tree(path), ast.Module)

    def test_parser_ownership_and_underlying_callable_boundaries_are_explicit(self) -> None:
        run_functions = _functions(RUN)
        a3_functions = _functions(A3)
        b1_functions = _functions(B1_RUNNER)
        compiler_functions = _functions(B1_COMPILER)
        b2_functions = _functions(B2)

        self.assertIn("parse_args", run_functions)
        self.assertIn("main", run_functions)
        self.assertIn("parse_cli_args", a3_functions)
        self.assertIn("main", a3_functions)
        self.assertIn("main", b1_functions)
        self.assertNotIn("parse_cli_args", b1_functions)
        self.assertIn("parse_cli_args = _impl.parse_cli_args", _source(B1_RUNNER))
        self.assertIn("def parse_cli_args(", _source(B1_IMPLEMENTATION))
        self.assertIn("def build_cli_parser(", _source(B2_IMPLEMENTATION))
        self.assertIn("def parse_arguments(", _source(B2_IMPLEMENTATION))
        self.assertIn("main", b2_functions)

        # The lower-level compiler deliberately owns neither a parser nor a callable
        # main boundary. Importing it executes compiler work, so it is not a safe
        # canonical dispatch target.
        self.assertNotIn("main", compiler_functions)
        self.assertNotIn("parse_args", compiler_functions)
        self.assertNotIn("parse_cli_args", compiler_functions)

        self.assertIn("args = parse_args()", _source(RUN))
        self.assertIn("_orchestrator.orchestrate_a3(", _source(A3))
        self.assertIn("_impl.orchestrate(", _source(B1_RUNNER))
        self.assertIn("b2_orchestrator.orchestrate_b2(", _source(B2))

    def test_direct_execution_guards_and_exception_translation_are_frozen(self) -> None:
        run_guard = _main_guard(RUN)
        a3_guard = _main_guard(A3)
        b1_guard = _main_guard(B1_RUNNER)
        b2_guard = _main_guard(B2)

        self.assertIsNotNone(run_guard)
        self.assertIsNotNone(a3_guard)
        self.assertIsNotNone(b1_guard)
        self.assertIsNotNone(b2_guard)
        self.assertIsNone(_main_guard(B1_COMPILER))

        run_guard_source = ast.get_source_segment(_source(RUN), run_guard)
        a3_guard_source = ast.get_source_segment(_source(A3), a3_guard)
        b1_guard_source = ast.get_source_segment(_source(B1_RUNNER), b1_guard)
        b2_guard_source = ast.get_source_segment(_source(B2), b2_guard)
        self.assertIsNotNone(run_guard_source)
        self.assertIsNotNone(a3_guard_source)
        self.assertIsNotNone(b1_guard_source)
        self.assertIsNotNone(b2_guard_source)

        self.assertIn("except subprocess.CalledProcessError as e:", run_guard_source)
        self.assertIn("sys.exit(e.returncode)", run_guard_source)
        self.assertIn("except Exception as e:", run_guard_source)
        self.assertIn("sys.exit(1)", run_guard_source)
        self.assertNotIn("KeyboardInterrupt", run_guard_source)
        self.assertIn("sys.exit(main())", a3_guard_source)
        self.assertIn("main()", b1_guard_source)
        self.assertNotIn("sys.exit(main())", b1_guard_source)
        self.assertIn("main()", b2_guard_source)
        self.assertNotIn("sys.exit(main())", b2_guard_source)

    def test_b1_compiler_remains_argumentless_and_import_executing(self) -> None:
        source = _source(B1_COMPILER)
        tree = _tree(B1_COMPILER)
        imported_modules = set()
        for node in tree.body:
            if isinstance(node, ast.Import):
                imported_modules.update(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                imported_modules.add(node.module)
        read_config_assignments = []
        for statement in tree.body:
            if not isinstance(statement, ast.Assign):
                continue
            if not any(
                isinstance(target, ast.Name) and target.id == "params"
                for target in statement.targets
            ):
                continue
            value = statement.value
            if (
                isinstance(value, ast.Call)
                and isinstance(value.func, ast.Attribute)
                and isinstance(value.func.value, ast.Name)
                and value.func.value.id == "_effects"
                and value.func.attr == "read_config"
            ):
                read_config_assignments.append(statement)

        self.assertNotIn("argparse", imported_modules)
        self.assertNotIn("sys.argv", source)
        self.assertEqual(len(read_config_assignments), 1)
        assignment_call = read_config_assignments[0].value
        self.assertIsInstance(assignment_call, ast.Call)
        self.assertEqual(len(assignment_call.args), 1)
        config_path = assignment_call.args[0]
        self.assertIsInstance(config_path, ast.Attribute)
        self.assertIsInstance(config_path.value, ast.Name)
        self.assertEqual(config_path.value.id, "_planning")
        self.assertEqual(config_path.attr, "CONFIG_PATH")
        self.assertEqual(assignment_call.keywords, [])
        self.assertIn("params = _effects.read_config(_planning.CONFIG_PATH)", source)

    def test_real_historical_help_and_malformed_parsers_stop_before_effects(
        self,
    ) -> None:
        cases = (
            (
                "run",
                importlib.import_module("run"),
                "check_tool_available",
                (("--help",), ("--env-name",), ("--unknown",)),
            ),
            (
                "A3",
                importlib.import_module("t1_confection.A3_process"),
                "orchestrate_a3",
                (("--help",), ("--scenario",), ("--unknown",)),
            ),
            (
                "B1",
                importlib.import_module("t1_confection.B1_Run_Compiler"),
                "orchestrate",
                (("--help",), ("--scenarios",), ("--unknown",)),
            ),
            (
                "B2",
                importlib.import_module("t1_confection.B2_Executing_OG_Model"),
                "_set_here",
                (("--help",), ("--scenarios",), ("--unknown",)),
            ),
        )

        for label, module, boundary_name, arguments in cases:
            if label == "A3":
                boundary_owner = module._orchestrator
            elif label == "B1":
                boundary_owner = module._impl
            else:
                boundary_owner = module

            for argv in arguments:
                with self.subTest(entrypoint=label, argv=argv):
                    boundary = mock.Mock(
                        side_effect=AssertionError(
                            f"{label} effect boundary must be unreachable"
                        )
                    )
                    with (
                        mock.patch.object(boundary_owner, boundary_name, boundary),
                        mock.patch.object(sys, "argv", [f"{label}.py", *argv]),
                    ):
                        code, _stdout, stderr = _capture_exit(module.main)

                    self.assertEqual(code, 0 if argv == ("--help",) else 2)
                    if code == 2:
                        self.assertIn("error:", stderr)
                    boundary.assert_not_called()

    def test_historical_path_and_process_semantics_remain_owned_by_entrypoints(self) -> None:
        run_source = _source(RUN)
        a3_source = _source(A3)
        b1_source = _source(B1_RUNNER)
        b1_impl_source = _source(B1_IMPLEMENTATION)
        b2_source = _source(B2)
        b2_impl_source = _source(B2_IMPLEMENTATION)

        # run.py is intentionally caller-CWD-relative and forwards inherited state
        # through its existing shell command boundary.
        self.assertIn('T1_DIR = Path("t1_confection")', run_source)
        self.assertIn('env = os.environ.copy()', run_source)
        self.assertIn('env["PYTHONHASHSEED"] = "0"', run_source)
        self.assertIn("shell=True, env=env", run_source)

        # A3 and B1 derive operational paths from their historical script files.
        self.assertIn("T1_CONFECTION = Path(__file__).resolve().parent", a3_source)
        self.assertIn("_impl.B1Paths.from_entrypoint(__file__)", b1_source)
        self.assertIn("cwd=str(command.cwd)", b1_impl_source)
        self.assertNotIn("shell=True", b1_impl_source)

        # B2 changes to its script-resolved HERE and owns both compile and optional
        # matrix/solver routes; it cannot truthfully be split by a friendly alias.
        self.assertIn('globals().get(', b2_source)
        self.assertIn("os.chdir(here)", b2_impl_source)
        self.assertIn("run_compiled_input_stage(", b2_impl_source)
        self.assertIn("run_execution_stage(", b2_impl_source)
        self.assertIn('params["execute_model"] or params["create_matrix"]', b2_impl_source)


if __name__ == "__main__":
    unittest.main()

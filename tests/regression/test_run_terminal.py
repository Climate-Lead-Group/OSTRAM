from __future__ import annotations

import io
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from unittest import mock

from ostram.terminal import RunReporter, safe_write


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]


class EncodingStream(io.StringIO):
    def __init__(self, encoding: str, *, tty: bool = False) -> None:
        super().__init__()
        self._encoding = encoding
        self._tty = tty

    @property
    def encoding(self) -> str:
        return self._encoding

    def isatty(self) -> bool:
        return self._tty

    def write(self, text: str) -> int:
        text.encode(self._encoding, errors="strict")
        return super().write(text)


class TerminalEncodingTests(unittest.TestCase):
    def test_cp1252_terminal_degrades_but_utf8_log_preserves_unicode_path(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp).resolve()
            workspace = root / "workspace ✓"
            stdout = EncodingStream("cp1252")
            stderr = EncodingStream("cp1252")
            reporter = RunReporter(
                project_root=REPO_ROOT,
                workspace=workspace,
                scenarios=("A_Calibrated_BAU",),
                verbose=False,
                stdout=stdout,
                stderr=stderr,
            )
            reporter.stage_start("B2", scenario="A_Calibrated_BAU")
            reporter.note(f"Reading {workspace / 'model ✓.txt'}")
            reporter.stage_complete("B2", detail="compile-only; solver skipped")
            reporter.finish(
                outcome="COMPILE_ONLY_SUCCESS",
                exit_code=0,
                final_message="OSTRAM compile-only run completed successfully",
            )

            terminal = stderr.getvalue()
            log = reporter.log_path.read_text(encoding="utf-8")

        self.assertIn(r"workspace \u2713", terminal)
        self.assertIn("workspace ✓", log)
        self.assertIn("model ✓.txt", log)
        self.assertIn("final_process_exit_code=0", log)

    def test_safe_write_handles_an_unencodable_value_directly(self) -> None:
        stream = EncodingStream("cp1252")
        safe_write(stream, "path ✓ / café", flush=True)
        self.assertEqual(stream.getvalue(), r"path \u2713 / café")


class RunReporterDisplayTests(unittest.TestCase):
    def test_interactive_status_updates_in_place_on_stderr_only(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            stdout = EncodingStream("utf-8", tty=True)
            stderr = EncodingStream("utf-8", tty=True)
            reporter = RunReporter(
                project_root=REPO_ROOT,
                workspace=Path(temp) / "workspace",
                scenarios=("BAU",),
                verbose=False,
                stdout=stdout,
                stderr=stderr,
            )
            reporter.stage_start("A1")
            reporter.recent_status("loading maintained inputs")
            reporter.stage_complete("A1")
            reporter.stage_skip("A2", "snapshot already exists")
            reporter.finish(
                outcome="SUCCESS",
                exit_code=0,
                final_message="OSTRAM run completed successfully",
            )
            terminal = stderr.getvalue()
            log = reporter.log_path.read_text(encoding="utf-8")

        self.assertIn("[1/5 · A1] Preparing the base model", terminal)
        self.assertIn("SKIPPED", terminal)
        self.assertIn("\x1b[2K", terminal)
        self.assertEqual(stdout.getvalue(), "")
        self.assertNotIn("\x1b", log)

    def test_noninteractive_status_is_append_only_and_has_no_ansi(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            stderr = EncodingStream("utf-8", tty=False)
            reporter = RunReporter(
                project_root=REPO_ROOT,
                workspace=Path(temp) / "workspace with spaces Ω",
                scenarios=("BAU",),
                verbose=False,
                stdout=EncodingStream("utf-8"),
                stderr=stderr,
            )
            reporter.stage_start("A3", scenario="BAU")
            reporter.stage_complete("A3")
            reporter.stage_skip("B1", "skipped by test")
            reporter.finish(
                outcome="SUCCESS",
                exit_code=0,
                final_message="OSTRAM run completed successfully",
            )
            terminal = stderr.getvalue()
            log = reporter.log_path.read_text(encoding="utf-8")

        self.assertIn("[3/5 · A3] Building selected scenarios | STARTED", terminal)
        self.assertIn("[4/5 · B1] Compiling model inputs | SKIPPED", terminal)
        self.assertNotIn("\x1b", terminal)
        self.assertNotIn("\r", terminal)
        self.assertNotIn("\x1b", log)


class ChildProcessTeeTests(unittest.TestCase):
    def _reporter(self, root: Path, *, verbose: bool):
        stdout = EncodingStream("utf-8")
        stderr = EncodingStream("utf-8")
        reporter = RunReporter(
            project_root=REPO_ROOT,
            workspace=root / "workspace child Ω",
            scenarios=("BAU",),
            verbose=verbose,
            stdout=stdout,
            stderr=stderr,
        )
        reporter.stage_start("B1", scenario="BAU")
        return reporter, stdout, stderr

    def test_concurrent_stdout_stderr_are_complete_without_deadlock(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp).resolve()
            reporter, stdout, stderr = self._reporter(root, verbose=False)
            script = (
                "import sys\n"
                "for i in range(1500):\n"
                " print(f'out-{i}-✓')\n"
                " print(f'err-{i}-Ω', file=sys.stderr)\n"
            )
            reporter.run_child(
                [sys.executable, "-B", "-c", script],
                cwd=root,
                env=os.environ.copy(),
            )
            reporter.stage_complete("B1")
            reporter.finish(
                outcome="SUCCESS",
                exit_code=0,
                final_message="OSTRAM run completed successfully",
            )
            log = reporter.log_path.read_text(encoding="utf-8")

        self.assertIn("out-0-✓", log)
        self.assertIn("out-1499-✓", log)
        self.assertIn("err-0-Ω", log)
        self.assertIn("err-1499-Ω", log)
        self.assertNotIn("out-0-✓", stdout.getvalue())
        self.assertNotIn("err-0-Ω", stderr.getvalue())

    def test_verbose_streams_raw_output_and_diagnostics_without_ansi(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp).resolve()
            reporter, stdout, stderr = self._reporter(root, verbose=True)
            reporter.run_child(
                [
                    sys.executable,
                    "-B",
                    "-c",
                    (
                        "import sys; "
                        "print('verbose stdout ✓'); "
                        "print('\\x1b[2Averbose stderr Ω', file=sys.stderr)"
                    ),
                ],
                cwd=root,
                env=os.environ.copy(),
            )
            reporter.stage_complete("B1")
            reporter.finish(
                outcome="SUCCESS",
                exit_code=0,
                final_message="OSTRAM run completed successfully",
            )
            log = reporter.log_path.read_text(encoding="utf-8")

        self.assertIn("verbose stdout ✓", stdout.getvalue())
        self.assertIn("verbose stderr Ω", stderr.getvalue())
        self.assertIn("Command:", stderr.getvalue())
        self.assertIn("Working directory:", stderr.getvalue())
        self.assertNotIn("\x1b", stdout.getvalue() + stderr.getvalue() + log)

    def test_nonzero_child_exit_code_is_preserved(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp).resolve()
            reporter, _stdout, _stderr = self._reporter(root, verbose=False)
            with self.assertRaises(subprocess.CalledProcessError) as raised:
                reporter.run_child(
                    [sys.executable, "-B", "-c", "raise SystemExit(7)"],
                    cwd=root,
                    env=os.environ.copy(),
                )
            reporter.finish(
                outcome="FAILED",
                exit_code=raised.exception.returncode,
                final_message="OSTRAM run failed",
            )
            log = reporter.log_path.read_text(encoding="utf-8")

        self.assertEqual(raised.exception.returncode, 7)
        self.assertIn("COMMAND | exit_code=7", log)
        self.assertIn("final_process_exit_code=7", log)

    def test_keyboard_interrupt_signals_child_and_propagates(self) -> None:
        class FakeProcess:
            def __init__(self) -> None:
                self.pid = 4242
                self.stdout = io.BytesIO()
                self.stderr = io.BytesIO()
                self.wait_calls = 0
                self.signal_received = None

            def wait(self, timeout=None):
                self.wait_calls += 1
                if self.wait_calls == 1:
                    raise KeyboardInterrupt()
                return 130

            def poll(self):
                return None

            def send_signal(self, sent_signal):
                self.signal_received = sent_signal

            def kill(self):
                raise AssertionError("graceful interrupt should have completed")

        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp).resolve()
            reporter, _stdout, _stderr = self._reporter(root, verbose=False)
            process = FakeProcess()
            with (
                mock.patch("ostram.terminal.subprocess.Popen", return_value=process),
                self.assertRaises(KeyboardInterrupt),
            ):
                reporter.run_child(
                    [sys.executable, "-B", "-c", "pass"],
                    cwd=root,
                    env=os.environ.copy(),
                )
            reporter.finish(
                outcome="INTERRUPTED",
                exit_code=130,
                final_message="OSTRAM run interrupted",
            )
            log = reporter.log_path.read_text(encoding="utf-8")

        self.assertIsNotNone(process.signal_received)
        self.assertIn("terminating child pid=4242", log)
        self.assertIn("final_process_exit_code=130", log)


class SideEffectFreeCliTests(unittest.TestCase):
    def test_help_malformed_import_and_inspection_do_not_create_logs(self) -> None:
        import ostram.__main__ as cli

        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp).resolve()
            workspaces = {
                name: root / name
                for name in ("help workspace", "bad workspace", "inspect workspace")
            }
            with (
                mock.patch.dict(os.environ, {"OSTRAM_WORKSPACE": str(workspaces["help workspace"])}),
                redirect_stdout(io.StringIO()),
                redirect_stderr(io.StringIO()),
            ):
                self.assertEqual(cli.main(["--help"]), 0)

            with (
                redirect_stdout(io.StringIO()),
                redirect_stderr(io.StringIO()),
                self.assertRaises(SystemExit) as raised,
            ):
                cli.main(
                    [
                        "--workspace",
                        str(workspaces["bad workspace"]),
                        "run",
                        "--not-an-option",
                    ]
                )
            self.assertEqual(raised.exception.code, 2)

            with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
                self.assertEqual(
                    cli.main(
                        [
                            "--project-root",
                            str(REPO_ROOT),
                            "--workspace",
                            str(workspaces["inspect workspace"]),
                            "inspect-resources",
                        ]
                    ),
                    0,
                )

            __import__("ostram.terminal")
            for workspace in workspaces.values():
                self.assertFalse(workspace.exists())


if __name__ == "__main__":
    unittest.main()

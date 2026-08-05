"""Terminal-safe progress, logging, and subprocess output for OSTRAM runs."""

from __future__ import annotations

from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from datetime import datetime, timezone
import io
import json
import locale
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
import threading
import time
from typing import Iterator, Sequence, TextIO


ANSI_ESCAPE = re.compile(r"\x1b(?:\[[0-?]*[ -/]*[@-~]|[@-_])")
SENSITIVE_ARGUMENT = re.compile(
    r"(?:token|password|passwd|secret|license|licence|license[-_]?server)",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class StageDefinition:
    number: int
    code: str
    description: str

    @property
    def label(self) -> str:
        return f"[{self.number}/5 · {self.code}] {self.description}"


STAGES = (
    StageDefinition(1, "A1", "Preparing the base model"),
    StageDefinition(2, "A2", "Adding the transmission network"),
    StageDefinition(3, "A3", "Building selected scenarios"),
    StageDefinition(4, "B1", "Compiling model inputs"),
    StageDefinition(5, "B2", "Running the model and collecting results"),
)


def _stream_encoding(stream: TextIO) -> str:
    return getattr(stream, "encoding", None) or locale.getpreferredencoding(False) or "utf-8"


def terminal_safe_text(text: object, stream: TextIO) -> str:
    """Return text representable by *stream* without changing the original value."""

    value = str(text)
    encoding = _stream_encoding(stream)
    try:
        value.encode(encoding, errors="strict")
        return value
    except (LookupError, UnicodeEncodeError):
        try:
            return value.encode(encoding, errors="backslashreplace").decode(encoding)
        except LookupError:
            return value.encode("ascii", errors="backslashreplace").decode("ascii")


def safe_write(stream: TextIO, text: object, *, flush: bool = False) -> int:
    """Write without allowing an incompatible terminal encoding to crash a run."""

    original = str(text)
    try:
        written = stream.write(original)
    except UnicodeEncodeError:
        written = stream.write(terminal_safe_text(original, stream))
    if flush:
        stream.flush()
    return len(original) if written is None else int(written)


def safe_print(
    *values: object,
    sep: str = " ",
    end: str = "\n",
    file: TextIO | None = None,
    flush: bool = False,
) -> None:
    """A small ``print`` equivalent that degrades unsupported glyphs safely."""

    stream = sys.stdout if file is None else file
    safe_write(stream, sep.join(str(value) for value in values) + end, flush=flush)


def _strip_terminal_controls(text: str) -> str:
    return ANSI_ESCAPE.sub("", text).replace("\r", "")


def _format_elapsed(seconds: float) -> str:
    hours, remainder = divmod(max(0, int(seconds)), 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def _redacted_command(tokens: Sequence[str]) -> str:
    redacted: list[str] = []
    hide_next = False
    for token in (str(item) for item in tokens):
        if hide_next:
            redacted.append("<redacted>")
            hide_next = False
            continue
        if "=" in token:
            name, value = token.split("=", 1)
            if SENSITIVE_ARGUMENT.search(name):
                redacted.append(f"{name}=<redacted>")
                continue
        redacted.append(token)
        if token.startswith("-") and SENSITIVE_ARGUMENT.search(token):
            hide_next = True
    return json.dumps(redacted, ensure_ascii=False)


def _decode_child_line(raw: bytes) -> str:
    for encoding in ("utf-8", locale.getpreferredencoding(False)):
        if not encoding:
            continue
        try:
            return raw.decode(encoding, errors="strict")
        except (LookupError, UnicodeDecodeError):
            pass
    return raw.decode("utf-8", errors="replace")


class _CapturedStream(io.TextIOBase):
    """Tee parent-process text into the UTF-8 run log and the real stream."""

    def __init__(self, reporter: "RunReporter", name: str, stream: TextIO) -> None:
        self.reporter = reporter
        self.name = name
        self.stream = stream
        self._buffer = ""

    @property
    def encoding(self) -> str:
        return _stream_encoding(self.stream)

    def isatty(self) -> bool:
        return bool(getattr(self.stream, "isatty", lambda: False)())

    def fileno(self) -> int:
        return self.stream.fileno()

    def writable(self) -> bool:
        return True

    def write(self, text: str) -> int:
        value = str(text)
        self._buffer += value.replace("\r", "\n")
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line:
                self.reporter.record_output(self.name, line)
        safe_write(self.stream, value)
        return len(value)

    def flush(self) -> None:
        if self._buffer:
            self.reporter.record_output(self.name, self._buffer)
            self._buffer = ""
        self.stream.flush()


class RunReporter:
    """Five-stage live status and detailed UTF-8 log for one actual run."""

    def __init__(
        self,
        *,
        project_root: Path,
        workspace: Path,
        scenarios: Sequence[str],
        verbose: bool,
        profile_id: str = "full",
        manifest: Path | None = None,
        compile_only: bool = False,
        stdout: TextIO | None = None,
        stderr: TextIO | None = None,
    ) -> None:
        self.project_root = Path(project_root).resolve()
        self.workspace = Path(workspace).resolve()
        self.scenarios = tuple(str(scenario) for scenario in scenarios)
        self.profile_id = str(profile_id)
        self.manifest = None if manifest is None else Path(manifest).resolve()
        self.compile_only = bool(compile_only)
        self.verbose = bool(verbose)
        self.stdout = sys.stdout if stdout is None else stdout
        self.stderr = sys.stderr if stderr is None else stderr
        self.interactive = bool(
            not self.verbose
            and getattr(self.stderr, "isatty", lambda: False)()
        )
        self.started_at = datetime.now(timezone.utc)
        self.started_monotonic = time.monotonic()
        run_id = self.started_at.strftime("%Y%m%dT%H%M%S%fZ")
        run_directory = self.workspace / "logs" / run_id
        suffix = 0
        while run_directory.exists():
            suffix += 1
            run_directory = self.workspace / "logs" / f"{run_id}-{suffix}"
        run_directory.mkdir(parents=True, exist_ok=False)
        self.log_path = run_directory / "run.log"
        self._log = self.log_path.open("x", encoding="utf-8", newline="\n")
        self._lock = threading.RLock()
        self._current: StageDefinition | None = None
        self._current_scenario: str | None = None
        self._stage_started = 0.0
        self._recent = "Initializing"
        self._states = {stage.code: "PENDING" for stage in STAGES}
        self._rendered = False
        self._finished = False
        self._capture_depth = 0
        self._stop = threading.Event()
        self._heartbeat: threading.Thread | None = None
        self._write_log("RUN", f"start_time={self.started_at.isoformat()}")
        self._write_log("RUN", f"project_root={self.project_root}")
        self._write_log("RUN", f"workspace={self.workspace}")
        self._write_log("RUN", f"profile_id={self.profile_id}")
        self._write_log("RUN", f"manifest={self.manifest}")
        self._write_log("RUN", "scenarios=" + json.dumps(self.scenarios, ensure_ascii=False))
        self._write_log("RUN", f"compile_only={self.compile_only}")
        self._write_log("RUN", f"verbose={self.verbose}")
        if self.interactive:
            self._heartbeat = threading.Thread(
                target=self._heartbeat_loop,
                name="ostram-run-status",
                daemon=True,
            )
            self._heartbeat.start()

    def _write_log(self, category: str, message: str) -> None:
        clean = _strip_terminal_controls(str(message))
        stamp = datetime.now(timezone.utc).isoformat()
        with self._lock:
            self._log.write(f"{stamp} | {category} | {clean}\n")
            self._log.flush()

    def record_output(self, stream_name: str, line: str) -> None:
        clean = _strip_terminal_controls(line)
        self._write_log(f"OUTPUT {stream_name}", clean)
        self.recent_status(clean)

    @contextmanager
    def capture_output(self) -> Iterator[None]:
        previous_stdout = sys.stdout
        previous_stderr = sys.stderr
        captured_stdout = _CapturedStream(self, "stdout", previous_stdout)
        captured_stderr = _CapturedStream(self, "stderr", previous_stderr)
        self._capture_depth += 1
        sys.stdout = captured_stdout
        sys.stderr = captured_stderr
        try:
            yield
        finally:
            captured_stdout.flush()
            captured_stderr.flush()
            sys.stdout = previous_stdout
            sys.stderr = previous_stderr
            self._capture_depth -= 1
            if self._finished and self._capture_depth == 0 and not self._log.closed:
                self._log.close()

    def _emit_status(self, text: str) -> None:
        if self.interactive:
            self._render()
        else:
            safe_write(self.stderr, text + "\n", flush=True)

    def _render(self) -> None:
        if not self.interactive or self._finished:
            return
        with self._lock:
            elapsed = _format_elapsed(time.monotonic() - self._stage_started)
            stage = self._current.label if self._current else "OSTRAM run setup"
            scenario = (
                f" | scenario {self._current_scenario}"
                if self._current_scenario
                else ""
            )
            state = self._states.get(self._current.code, "RUNNING") if self._current else "RUNNING"
            line = f"{stage} | {state}{scenario} | elapsed {elapsed} | {self._recent}"
            safe_write(self.stderr, "\r\x1b[2K" + line[:180], flush=True)
            self._rendered = True

    def _heartbeat_loop(self) -> None:
        while not self._stop.wait(0.5):
            self._render()

    def recent_status(self, message: str) -> None:
        clean = " ".join(str(message).split())
        if not clean:
            return
        with self._lock:
            self._recent = clean[:140]
            for scenario in self.scenarios:
                if scenario in clean:
                    self._current_scenario = scenario
                    break
        self._render()

    def stage_start(self, code: str, *, scenario: str | None = None) -> None:
        stage = next(item for item in STAGES if item.code == code)
        with self._lock:
            self._current = stage
            self._current_scenario = scenario
            self._stage_started = time.monotonic()
            self._states[code] = "RUNNING"
            self._recent = "Started"
        self._write_log("STAGE", f"{stage.code} STARTED | {stage.description}")
        self._emit_status(f"{stage.label} | STARTED")

    def stage_complete(self, code: str, *, detail: str | None = None) -> None:
        stage = next(item for item in STAGES if item.code == code)
        elapsed = time.monotonic() - self._stage_started
        with self._lock:
            self._states[code] = "COMPLETED"
            self._recent = detail or "Completed"
        suffix = f" | {detail}" if detail else ""
        self._write_log(
            "STAGE",
            f"{stage.code} COMPLETED | elapsed={_format_elapsed(elapsed)}{suffix}",
        )
        self._emit_status(
            f"{stage.label} | COMPLETED | elapsed {_format_elapsed(elapsed)}{suffix}"
        )

    def stage_skip(self, code: str, reason: str) -> None:
        stage = next(item for item in STAGES if item.code == code)
        with self._lock:
            self._current = stage
            self._current_scenario = None
            self._stage_started = time.monotonic()
            self._states[code] = "SKIPPED"
            self._recent = reason
        self._write_log("STAGE", f"{stage.code} SKIPPED | {reason}")
        self._emit_status(f"{stage.label} | SKIPPED | {reason}")

    def stage_fail(self, detail: str) -> None:
        with self._lock:
            stage = self._current
            if stage is not None:
                self._states[stage.code] = "FAILED"
            self._recent = detail
        if stage is not None:
            self._write_log("STAGE", f"{stage.code} FAILED | {detail}")
            self._emit_status(f"{stage.label} | FAILED | {detail}")

    def note(self, message: str) -> None:
        self._write_log("STATUS", message)
        self.recent_status(message)

    def _log_command(self, command: Sequence[str], cwd: Path | None) -> None:
        self._write_log("COMMAND", f"tokens={_redacted_command(command)}")
        self._write_log("COMMAND", f"cwd={Path.cwd() if cwd is None else Path(cwd).resolve()}")

    def _read_pipe(self, pipe: io.BufferedReader, name: str) -> None:
        try:
            for raw in iter(pipe.readline, b""):
                text = _decode_child_line(raw).rstrip("\r\n")
                clean = _strip_terminal_controls(text)
                self._write_log(f"CHILD {name}", clean)
                self.recent_status(clean)
                if self.verbose:
                    target = self.stdout if name == "stdout" else self.stderr
                    safe_write(target, clean + "\n", flush=True)
        finally:
            pipe.close()

    def _interrupt_child(self, process: subprocess.Popen[bytes]) -> None:
        self._write_log("INTERRUPT", f"terminating child pid={process.pid}")
        if process.poll() is not None:
            return
        if os.name == "nt":
            try:
                process.send_signal(signal.CTRL_BREAK_EVENT)
                process.wait(timeout=2)
                return
            except (OSError, subprocess.TimeoutExpired):
                subprocess.run(
                    ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                    shell=False,
                )
        else:
            try:
                os.killpg(process.pid, signal.SIGINT)
                process.wait(timeout=2)
                return
            except (OSError, subprocess.TimeoutExpired):
                try:
                    os.killpg(process.pid, signal.SIGTERM)
                except OSError:
                    pass
        try:
            process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()

    def run_child(
        self,
        command: Sequence[str],
        *,
        cwd: Path | None,
        env: dict[str, str],
    ) -> None:
        tokens = [str(token) for token in command]
        child_env = dict(env)
        child_env["PYTHONIOENCODING"] = "utf-8"
        child_env.setdefault("NO_COLOR", "1")
        self._log_command(tokens, cwd)
        if self.verbose:
            safe_write(self.stderr, f"Command: {_redacted_command(tokens)}\n", flush=True)
            safe_write(
                self.stderr,
                f"Working directory: {Path.cwd() if cwd is None else Path(cwd).resolve()}\n",
                flush=True,
            )
        creationflags = 0
        popen_kwargs: dict[str, object] = {}
        if os.name == "nt":
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP
        else:
            popen_kwargs["start_new_session"] = True
        process = subprocess.Popen(
            tokens,
            cwd=str(Path(cwd).resolve()) if cwd is not None else None,
            env=child_env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
            creationflags=creationflags,
            **popen_kwargs,
        )
        assert process.stdout is not None
        assert process.stderr is not None
        readers = [
            threading.Thread(
                target=self._read_pipe,
                args=(process.stdout, "stdout"),
                name="ostram-child-stdout",
            ),
            threading.Thread(
                target=self._read_pipe,
                args=(process.stderr, "stderr"),
                name="ostram-child-stderr",
            ),
        ]
        for reader in readers:
            reader.start()
        try:
            returncode = process.wait()
        except KeyboardInterrupt:
            self._interrupt_child(process)
            raise
        finally:
            for reader in readers:
                reader.join()
        self._write_log("COMMAND", f"exit_code={returncode}")
        if returncode:
            raise subprocess.CalledProcessError(returncode, tokens)

    def finish(self, *, outcome: str, exit_code: int, final_message: str) -> None:
        if self._finished:
            return
        self._stop.set()
        if self._heartbeat is not None:
            self._heartbeat.join(timeout=1)
        elapsed = time.monotonic() - self.started_monotonic
        ended_at = datetime.now(timezone.utc)
        self._write_log("RUN", f"end_time={ended_at.isoformat()}")
        self._write_log("RUN", f"outcome={outcome}")
        self._write_log("RUN", f"elapsed={_format_elapsed(elapsed)}")
        self._write_log("RUN", f"final_process_exit_code={exit_code}")
        self._finished = True
        if self.interactive and self._rendered:
            safe_write(self.stderr, "\r\x1b[2K", flush=True)
        safe_write(self.stderr, final_message + "\n", flush=True)
        safe_write(self.stderr, f"Detailed log: {self.log_path}\n", flush=True)
        if self._capture_depth == 0:
            self._log.close()


_ACTIVE_REPORTER: ContextVar[RunReporter | None] = ContextVar(
    "ostram_active_run_reporter", default=None
)


@contextmanager
def activate_reporter(reporter: RunReporter) -> Iterator[None]:
    token = _ACTIVE_REPORTER.set(reporter)
    try:
        yield
    finally:
        _ACTIVE_REPORTER.reset(token)


def active_reporter() -> RunReporter | None:
    return _ACTIVE_REPORTER.get()

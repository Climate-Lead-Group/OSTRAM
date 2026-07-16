#!/usr/bin/env python3
"""Offline regression evidence for OSTRAM.

This module is intentionally solver-free.  Its only child process is ``git`` for
read-only repository metadata.  It discovers scenario snapshots, normalizes and
hashes static/generated artifacts, captures compact manifests, and compares two
captured evidence directories.  It never calls the OSTRAM pipeline, DVC, GLPK,
CPLEX, or another optimizer.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import math
import os
import platform
import re
import subprocess
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator, Mapping, Sequence


HERE = Path(__file__).resolve().parent
DEFAULT_SCENARIOS = HERE / "scenarios.yaml"
DEFAULT_TOLERANCES = HERE / "tolerances.yaml"
ZERO_THRESHOLD = 1e-6
INTEGER_LIKE_COLUMNS = {"YEAR", "MODE_OF_OPERATION"}
INDEX_COLUMNS = {"Unnamed: 0"}
PRESERVATION_SCOPES = {"preservation", "regression"}
CLEANUP_ACCEPTANCE_SCOPE = "cleanup-acceptance"
STATIC_ACCEPTANCE_STAGES = ("a1", "config", "a2", "otoole")

PROTECTED_TREE_ROOTS = (
    "t1_confection/A1_Outputs",
    "t1_confection/A2_Output_Params",
    "t1_confection/A2_Outputs_Params_otoole",
    "t1_confection/OG_csvs_inputs",
    "t1_confection/A2_Extra_Inputs",
    "t1_confection/Miscellaneous",
    "t1_confection/templates",
    "t1_confection/A3_process/rules_scripts",
    "t1_confection/sensitivity_expansion/reference",
)
PROTECTED_FILES = (
    ".gitattributes",
    "run.py",
    "dvc.yaml",
    "dvc.lock",
    "t1_confection/A1_Pre_processing_OG_csvs.py",
    "t1_confection/A2_AddTx.py",
    "t1_confection/A3_process.py",
    "t1_confection/B1_Run_Compiler.py",
    "t1_confection/b1_runner.py",
    "t1_confection/B1_Compiler.py",
    "t1_confection/B2_Executing_OG_Model.py",
    "t1_confection/b2_orchestrator.py",
    "t1_confection/Config_MOMF_T1_A.yaml",
    "t1_confection/Config_MOMF_T1_AB.yaml",
    "t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx",
    "t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v17.xlsx",
    "t1_confection/A3_process/Interconnectors.xlsx",
    "t1_confection/A3_process/TECH_TYPES.csv",
)
PROTECTED_GLOBS = (
    "t1_confection/osemosys_fast_preprocessed*.txt",
    "t1_confection/A3_process/*.py",
    "t1_confection/sensitivity_expansion/*.py",
)


class RegressionError(ValueError):
    """Raised when evidence is structurally invalid."""


@dataclass(frozen=True)
class NormalizedCSV:
    columns: tuple[str, ...]
    key_columns: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]
    payload: bytes


@dataclass(frozen=True)
class Comparison:
    status: str
    details: tuple[str, ...] = ()

    @property
    def passed(self) -> bool:
        return self.status in {"exact", "normalized-exact", "numeric-equivalent/hash-drift"}


def load_json_yaml(path: Path) -> dict:
    """Load the repository's JSON-compatible YAML without a YAML dependency."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RegressionError(f"cannot load {path}: {exc}") from exc
    if not isinstance(data, dict):
        raise RegressionError(f"{path} must contain a mapping")
    return data


def load_scenarios(path: Path = DEFAULT_SCENARIOS) -> list[dict]:
    data = load_json_yaml(path)
    scenarios = data.get("scenarios")
    if not isinstance(scenarios, list) or not scenarios:
        raise RegressionError("scenario inventory is empty or invalid")
    names = [item.get("name") for item in scenarios if isinstance(item, dict)]
    if len(names) != len(scenarios) or any(not isinstance(name, str) or not name for name in names):
        raise RegressionError("every scenario must have a non-empty string name")
    duplicates = sorted(name for name, count in Counter(names).items() if count > 1)
    if duplicates:
        raise RegressionError(f"duplicate scenario names: {', '.join(duplicates)}")
    return scenarios


def scenarios_for_scope(
    inventory: Sequence[Mapping[str, object]],
    scope: str,
) -> list[Mapping[str, object]]:
    """Return an explicit policy scope without dropping preservation inventory entries."""
    if scope in PRESERVATION_SCOPES:
        return list(inventory)
    if scope != CLEANUP_ACCEPTANCE_SCOPE:
        raise RegressionError(f"unknown scenario scope: {scope}")
    if any(not isinstance(item.get("cleanup_acceptance"), bool) for item in inventory):
        raise RegressionError("every scenario must declare boolean cleanup_acceptance")
    excluded = [item for item in inventory if not item["cleanup_acceptance"]]
    missing_reasons = [
        str(item["name"])
        for item in excluded
        if not isinstance(item.get("cleanup_exclusion_reason"), str)
        or not str(item["cleanup_exclusion_reason"]).strip()
    ]
    if missing_reasons:
        raise RegressionError(
            "cleanup exclusions require reasons: " + ", ".join(sorted(missing_reasons))
        )
    selected = [item for item in inventory if item["cleanup_acceptance"]]
    if len(inventory) != 20 or len(selected) != 16 or len(excluded) != 4:
        raise RegressionError("cleanup policy must preserve 20 scenarios and accept exactly 16")
    return selected


def _scenario_dirs(parent: Path, prefix: str = "") -> set[str]:
    if not parent.is_dir():
        return set()
    result = set()
    for child in parent.iterdir():
        if not child.is_dir() or (prefix and not child.name.startswith(prefix)):
            continue
        name = child.name[len(prefix):] if prefix else child.name
        if name and not name.startswith("_"):
            result.add(name)
    return result


def discover_scenarios(repo: Path, inventory: Sequence[Mapping[str, object]]) -> dict:
    repo = repo.resolve()
    expected = {str(item["name"]) for item in inventory}
    a1 = _scenario_dirs(repo / "t1_confection" / "A1_Outputs", "A1_Outputs_")
    configs = _scenario_dirs(repo / "t1_confection" / "A3_process" / "rules_scripts" / "configs")
    a2 = _scenario_dirs(repo / "t1_confection" / "A2_Output_Params")
    otoole = _scenario_dirs(repo / "t1_confection" / "A2_Outputs_Params_otoole")
    return {
        "expected": expected,
        "a1": a1,
        "configs": configs,
        "a2": a2,
        "otoole": otoole,
        "missing_a1": expected - a1,
        "unexpected_a1": a1 - expected,
        "missing_configs": expected - configs,
        "unexpected_configs": configs - expected,
        "missing_a2": expected - a2,
        "missing_otoole": expected - otoole,
    }


def discovery_passes(discovery: Mapping[str, set[str]]) -> bool:
    expected = discovery["expected"]
    return len(expected) == 20 and discovery["a1"] == expected and discovery["configs"] == expected


def cleanup_acceptance_discovery_passes(discovery: Mapping[str, set[str]]) -> bool:
    expected = discovery["expected"]
    return (
        len(expected) == 16
        and expected <= discovery["a1"]
        and expected <= discovery["configs"]
        and expected <= discovery["a2"]
        and expected <= discovery["otoole"]
    )


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def _normal_text(value: object) -> str:
    return unicodedata.normalize("NFKC", "" if value is None else str(value)).strip()


def _normal_integer_like(value: str, column: str) -> str:
    if column not in INTEGER_LIKE_COLUMNS or not value:
        return value
    try:
        number = float(value)
    except ValueError:
        return value
    if math.isfinite(number) and number.is_integer():
        return str(int(number))
    return value


def _normal_number(value: str, zero_threshold: float) -> str:
    try:
        number = float(value)
    except ValueError as exc:
        raise RegressionError(f"invalid VALUE {value!r}") from exc
    if not math.isfinite(number):
        raise RegressionError(f"non-finite VALUE {value!r}")
    if abs(number) <= zero_threshold:
        number = 0.0
    return format(number, ".15g")


def normalize_csv_text(
    text: str,
    *,
    required_columns: Iterable[str] = (),
    zero_threshold: float = ZERO_THRESHOLD,
    omit_blank_values: bool = False,
) -> NormalizedCSV:
    reader = csv.DictReader(io.StringIO(text, newline=""))
    if reader.fieldnames is None:
        raise RegressionError("CSV has no header")
    original_columns = [_normal_text(name) for name in reader.fieldnames]
    if len(original_columns) != len(set(original_columns)):
        raise RegressionError("CSV has duplicate column names")
    columns = tuple(name for name in original_columns if name not in INDEX_COLUMNS)
    missing = sorted(set(required_columns) - set(columns))
    if missing:
        raise RegressionError(f"missing required columns: {', '.join(missing)}")
    value_column = next((name for name in columns if name.upper() == "VALUE"), None)
    # OSeMOSYS set tables such as DAILYTIMEBRACKET.csv contain a single column
    # named VALUE.  In that format VALUE is the set member (and therefore the
    # key), not a numeric measure.
    if len(columns) == 1:
        value_column = None
    key_columns = tuple(name for name in columns if name != value_column)
    rows: list[tuple[str, ...]] = []
    keys: set[tuple[str, ...]] = set()
    for line_number, raw in enumerate(reader, start=2):
        normalized: dict[str, str] = {}
        for original, column in zip(reader.fieldnames, original_columns):
            if column in INDEX_COLUMNS:
                continue
            value = _normal_text(raw.get(original))
            value = _normal_integer_like(value, column.upper())
            normalized[column] = value
        # Spreadsheet-derived CSVs pad shorter set columns with fully blank rows.
        # They contain no key/value information and are safe to omit.
        if all(not value for value in normalized.values()):
            continue
        if value_column is not None:
            if not normalized[value_column] and omit_blank_values:
                continue
            normalized[value_column] = _normal_number(normalized[value_column], zero_threshold)
        key = tuple(normalized[column] for column in key_columns)
        if key in keys:
            raise RegressionError(f"duplicate key at line {line_number}: {key!r}")
        keys.add(key)
        rows.append(tuple(normalized[column] for column in columns))
    key_indexes = tuple(columns.index(column) for column in key_columns)
    rows.sort(key=lambda row: tuple(row[index] for index in key_indexes))
    out = io.StringIO(newline="")
    writer = csv.writer(out, lineterminator="\n")
    writer.writerow(columns)
    writer.writerows(rows)
    return NormalizedCSV(columns, key_columns, tuple(rows), out.getvalue().encode("utf-8"))


def normalize_csv_file(path: Path, **kwargs: object) -> NormalizedCSV:
    return normalize_csv_text(path.read_text(encoding="utf-8-sig"), **kwargs)


def normalize_text_bytes(raw: bytes) -> bytes:
    text = raw.decode("utf-8-sig", errors="strict")
    text = unicodedata.normalize("NFC", text).replace("\r\n", "\n").replace("\r", "\n")
    return (text.rstrip("\n") + "\n").encode("utf-8")


_CORE_DATE = re.compile(
    rb"(<dcterms:(?:created|modified)\b[^>]*>).*?(</dcterms:(?:created|modified)>)",
    flags=re.DOTALL,
)

_XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _xlsx_cell_text(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    kind = cell.get("t")
    value = cell.find(f"{{{_XLSX_NS}}}v")
    if kind == "s" and value is not None and value.text:
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return ""
    if kind == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(f".//{{{_XLSX_NS}}}t"))
    return value.text if value is not None and value.text else ""


def _column_letters(reference: str) -> str:
    return "".join(char for char in reference if char.isalpha()).upper()


def _normalize_restrictions_timestamp(members: dict[str, bytes]) -> None:
    workbook_data = members.get("xl/workbook.xml")
    rels_data = members.get("xl/_rels/workbook.xml.rels")
    if workbook_data is None or rels_data is None:
        return
    try:
        workbook = ET.fromstring(workbook_data)
        rels = ET.fromstring(rels_data)
    except ET.ParseError:
        return
    rel_id = None
    for sheet in workbook.findall(f".//{{{_XLSX_NS}}}sheet"):
        if (sheet.get("name") or "").strip().casefold() == "restrictions":
            rel_id = sheet.get(f"{{{_DOC_REL_NS}}}id")
            break
    if not rel_id:
        return
    target = None
    for relation in rels.findall(f"{{{_PKG_REL_NS}}}Relationship"):
        if relation.get("Id") == rel_id:
            target = relation.get("Target")
            break
    if not target:
        return
    sheet_name = target.lstrip("/")
    if not sheet_name.startswith("xl/"):
        sheet_name = "xl/" + sheet_name
    sheet_name = str(Path(sheet_name).as_posix())
    sheet_data = members.get(sheet_name)
    if sheet_data is None:
        return
    shared_strings: list[str] = []
    shared_data = members.get("xl/sharedStrings.xml")
    if shared_data:
        try:
            shared = ET.fromstring(shared_data)
            shared_strings = [
                "".join(node.text or "" for node in item.findall(f".//{{{_XLSX_NS}}}t"))
                for item in shared.findall(f"{{{_XLSX_NS}}}si")
            ]
        except ET.ParseError:
            shared_strings = []
    try:
        sheet = ET.fromstring(sheet_data)
    except ET.ParseError:
        return
    header_column = None
    rows = sheet.findall(f".//{{{_XLSX_NS}}}row")
    for row in rows:
        for cell in row.findall(f"{{{_XLSX_NS}}}c"):
            if _normal_text(_xlsx_cell_text(cell, shared_strings)).casefold() == "source_run_timestamp":
                header_column = _column_letters(cell.get("r") or "")
                break
        if header_column:
            break
    if not header_column:
        return
    for row in rows:
        for cell in row.findall(f"{{{_XLSX_NS}}}c"):
            if _column_letters(cell.get("r") or "") != header_column:
                continue
            if _normal_text(_xlsx_cell_text(cell, shared_strings)).casefold() == "source_run_timestamp":
                continue
            for child in list(cell):
                cell.remove(child)
            cell.set("t", "inlineStr")
            inline = ET.SubElement(cell, f"{{{_XLSX_NS}}}is")
            text_node = ET.SubElement(inline, f"{{{_XLSX_NS}}}t")
            text_node.text = "NORMALIZED"
    members[sheet_name] = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)


def normalize_xlsx_members(members: Mapping[str, bytes]) -> dict[str, bytes]:
    normalized = dict(members)
    core = normalized.get("docProps/core.xml")
    if core is not None:
        normalized["docProps/core.xml"] = _CORE_DATE.sub(rb"\1NORMALIZED\2", core)
    _normalize_restrictions_timestamp(normalized)
    return normalized


def normalized_xlsx_sha256(path: Path) -> str:
    """Hash XLSX contents without container/core/Restrictions timestamps."""
    digest = hashlib.sha256()
    try:
        with zipfile.ZipFile(path) as archive:
            members = normalize_xlsx_members({
                info.filename: archive.read(info.filename)
                for info in archive.infolist() if not info.is_dir()
            })
            for name, payload in sorted(members.items()):
                encoded_name = unicodedata.normalize("NFC", name).encode("utf-8")
                digest.update(len(encoded_name).to_bytes(4, "big"))
                digest.update(encoded_name)
                digest.update(len(payload).to_bytes(8, "big"))
                digest.update(payload)
    except zipfile.BadZipFile as exc:
        raise RegressionError(f"invalid XLSX file {path}") from exc
    return digest.hexdigest()


def normalized_file_hash(path: Path, *, omit_blank_values: bool = False) -> tuple[str, str]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        payload = normalize_csv_file(path, omit_blank_values=omit_blank_values).payload
        policy = "csv-v1-blank-values-omitted" if omit_blank_values else "csv-v1"
        return hashlib.sha256(payload).hexdigest(), policy
    if suffix == ".xlsx":
        return normalized_xlsx_sha256(path), "xlsx-members-v1"
    if suffix in {".txt", ".yaml", ".yml", ".json", ".py", ".md", ".rst", ".bat", ".log"}:
        payload = normalize_text_bytes(path.read_bytes())
        return hashlib.sha256(payload).hexdigest(), "text-lf-nfc-v1"
    return sha256_file(path), "raw-v1"


def _git(repo: Path, *args: str) -> str | None:
    try:
        completed = subprocess.run(
            ["git", "-C", str(repo), *args],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
    except OSError:
        return None
    # Preserve leading spaces: porcelain status uses them as the first XY flag.
    # Only line terminators at the end are transport noise.
    return completed.stdout.rstrip("\r\n") if completed.returncode == 0 else None


def parse_porcelain_paths(status: str | None) -> list[str]:
    if not status:
        return []
    paths = []
    for line in status.splitlines():
        if len(line) >= 4:
            paths.append(line[3:])
    return sorted(paths)


def git_metadata(repo: Path) -> dict:
    status = _git(repo, "status", "--porcelain=v1")
    remote = _git(repo, "remote", "get-url", "origin")
    return {
        "commit": _git(repo, "rev-parse", "HEAD"),
        "tree": _git(repo, "show", "-s", "--format=%T", "HEAD"),
        "branch": _git(repo, "branch", "--show-current"),
        "remote": remote,
        "dirty_paths": parse_porcelain_paths(status),
    }


def tracked_files(repo: Path) -> set[str] | None:
    output = _git(repo, "ls-files", "-z")
    if output is None:
        return None
    return {item.replace("\\", "/") for item in output.split("\0") if item}


def _tree_has_tracked(repo: Path, path: Path, tracked: set[str] | None) -> bool:
    if tracked is None:
        return path.exists()
    try:
        rel = path.relative_to(repo).as_posix()
    except ValueError:
        return False
    prefix = rel.rstrip("/") + "/"
    return rel in tracked or any(item.startswith(prefix) for item in tracked)


def compiled_text_files(repo: Path, scenario: str) -> list[Path]:
    folder = repo / "t1_confection" / "Executables" / f"{scenario}_0"
    if not folder.is_dir():
        return []
    return sorted(
        path for path in folder.glob(f"Pre_processed_{scenario}_0*.txt")
        if path.is_file() and not path.name.endswith(".warnings.txt")
    )


def output_csv_files(repo: Path, scenario: str) -> list[Path]:
    folder = repo / "t1_confection" / "Executables" / f"{scenario}_0" / "Outputs"
    return sorted(folder.glob("*.csv")) if folder.is_dir() else []


def scenario_coverage(
    repo: Path,
    scenario: str,
    *,
    tracked_only: bool,
    tracked: set[str] | None = None,
) -> dict[str, object]:
    paths = {
        "a1": repo / "t1_confection" / "A1_Outputs" / f"A1_Outputs_{scenario}",
        "config": repo / "t1_confection" / "A3_process" / "rules_scripts" / "configs" / scenario,
        "a2": repo / "t1_confection" / "A2_Output_Params" / scenario,
        "otoole": repo / "t1_confection" / "A2_Outputs_Params_otoole" / scenario,
    }
    result: dict[str, object] = {}
    for stage, path in paths.items():
        present = _tree_has_tracked(repo, path, tracked) if tracked_only else path.is_dir()
        result[stage] = present
    compiled = compiled_text_files(repo, scenario)
    outputs = output_csv_files(repo, scenario)
    if tracked_only and tracked is not None:
        compiled = [path for path in compiled if path.relative_to(repo).as_posix() in tracked]
        outputs = [path for path in outputs if path.relative_to(repo).as_posix() in tracked]
    result["compiled_txt_count"] = len(compiled)
    result["output_csv_count"] = len(outputs)
    return result


def _iter_files(path: Path) -> Iterator[Path]:
    if path.is_file():
        yield path
    elif path.is_dir():
        yield from sorted(item for item in path.rglob("*") if item.is_file())


def artifact_files(repo: Path, scenario: str, *, include_generated: bool) -> Iterator[tuple[str, Path]]:
    stage_roots = {
        "a1": repo / "t1_confection" / "A1_Outputs" / f"A1_Outputs_{scenario}",
        "config": repo / "t1_confection" / "A3_process" / "rules_scripts" / "configs" / scenario,
        "a2": repo / "t1_confection" / "A2_Output_Params" / scenario,
        "otoole": repo / "t1_confection" / "A2_Outputs_Params_otoole" / scenario,
    }
    for stage, root in stage_roots.items():
        for path in _iter_files(root):
            yield stage, path
    if include_generated:
        for path in compiled_text_files(repo, scenario):
            yield "compiled", path
        for path in output_csv_files(repo, scenario):
            yield "outputs", path


def _record_for_file(source: str, scenario: str, stage: str, repo: Path, path: Path) -> dict:
    try:
        normalized, policy = normalized_file_hash(path, omit_blank_values=(stage in {"a2", "otoole"}))
    except RegressionError as exc:
        raise RegressionError(f"{path.relative_to(repo).as_posix()}: {exc}") from exc
    return {
        "source": source,
        "scenario": scenario,
        "stage": stage,
        "path": path.relative_to(repo).as_posix(),
        "size": path.stat().st_size,
        "raw_sha256": sha256_file(path),
        "normalized_sha256": normalized,
        "normalization": policy,
    }


def collect_artifact_records(
    source: str,
    repo: Path,
    scenarios: Sequence[str],
    *,
    tracked_only: bool,
    include_generated: bool,
) -> list[dict]:
    tracked = tracked_files(repo) if tracked_only else None
    records: list[dict] = []
    for scenario in scenarios:
        for stage, path in artifact_files(repo, scenario, include_generated=include_generated):
            rel = path.relative_to(repo).as_posix()
            if tracked_only and tracked is not None and rel not in tracked:
                continue
            records.append(_record_for_file(source, scenario, stage, repo, path))
    return records


def aggregate_records(records: Sequence[Mapping[str, object]]) -> list[dict]:
    groups: dict[tuple[str, str, str], list[Mapping[str, object]]] = {}
    for record in records:
        key = (str(record["source"]), str(record["scenario"]), str(record["stage"]))
        groups.setdefault(key, []).append(record)
    metrics = []
    for (source, scenario, stage), items in sorted(groups.items()):
        digest = hashlib.sha256()
        for item in sorted(items, key=lambda value: str(value["path"])):
            digest.update(str(item["path"]).encode("utf-8"))
            digest.update(b"\0")
            digest.update(str(item["normalized_sha256"]).encode("ascii"))
            digest.update(b"\n")
        metrics.append({
            "source": source,
            "scenario": scenario,
            "stage": stage,
            "file_count": len(items),
            "total_bytes": sum(int(item["size"]) for item in items),
            "aggregate_normalized_sha256": digest.hexdigest(),
        })
    return metrics


def _protected_file_list(repo: Path) -> list[Path]:
    found: set[Path] = set()
    for rel in PROTECTED_TREE_ROOTS:
        found.update(_iter_files(repo / rel))
    for rel in PROTECTED_FILES:
        path = repo / rel
        if path.is_file():
            found.add(path)
    for pattern in PROTECTED_GLOBS:
        found.update(path for path in repo.glob(pattern) if path.is_file())
    return sorted(found)


def protected_snapshot(repo: Path) -> dict:
    digest = hashlib.sha256()
    count = 0
    total_bytes = 0
    for path in _protected_file_list(repo):
        rel = path.relative_to(repo).as_posix()
        file_hash = sha256_file(path)
        digest.update(rel.encode("utf-8"))
        digest.update(b"\0")
        digest.update(file_hash.encode("ascii"))
        digest.update(b"\n")
        count += 1
        total_bytes += path.stat().st_size
    return {"file_count": count, "total_bytes": total_bytes, "aggregate_raw_sha256": digest.hexdigest()}


def required_files_report(root: Path, relative_paths: Iterable[str]) -> dict:
    missing = sorted(rel for rel in relative_paths if not (root / rel).is_file())
    return {"ok": not missing, "missing": missing}


def compare_csv_files(
    baseline: Path,
    candidate: Path,
    *,
    absolute: float = 1e-6,
    relative: float = 1e-8,
) -> Comparison:
    if not baseline.is_file() or not candidate.is_file():
        missing = [str(path) for path in (baseline, candidate) if not path.is_file()]
        return Comparison("missing-file", tuple(missing))
    if sha256_file(baseline) == sha256_file(candidate):
        return Comparison("exact")
    try:
        left = normalize_csv_file(baseline)
        right = normalize_csv_file(candidate)
    except RegressionError as exc:
        return Comparison("invalid", (str(exc),))
    if left.columns != right.columns or left.key_columns != right.key_columns:
        return Comparison("schema-drift", (f"{left.columns!r} != {right.columns!r}",))
    if left.payload == right.payload:
        return Comparison("normalized-exact")
    value_index = next((i for i, name in enumerate(left.columns) if name.upper() == "VALUE"), None)
    key_indexes = tuple(left.columns.index(name) for name in left.key_columns)
    left_map = {tuple(row[i] for i in key_indexes): row for row in left.rows}
    right_map = {tuple(row[i] for i in key_indexes): row for row in right.rows}
    if left_map.keys() != right_map.keys():
        return Comparison("key-drift", (f"missing={len(left_map.keys()-right_map.keys())}", f"extra={len(right_map.keys()-left_map.keys())}"))
    if value_index is None:
        return Comparison("value-drift", ("no VALUE column for tolerance comparison",))
    failures = []
    for key in sorted(left_map):
        a = float(left_map[key][value_index])
        b = float(right_map[key][value_index])
        if not math.isclose(a, b, rel_tol=relative, abs_tol=absolute):
            failures.append(f"{key!r}: {a} != {b}")
            if len(failures) == 10:
                break
    return Comparison("numeric-drift" if failures else "numeric-equivalent/hash-drift", tuple(failures))


def compare_hash_records(baseline: Sequence[Mapping[str, object]], candidate: Sequence[Mapping[str, object]]) -> dict:
    def key(record: Mapping[str, object]) -> tuple[str, str, str, str]:
        return (
            str(record.get("source", "")),
            str(record["scenario"]),
            str(record["stage"]),
            str(record["path"]),
        )
    left = {key(item): item for item in baseline}
    right = {key(item): item for item in candidate}
    missing = sorted(left.keys() - right.keys())
    extra = sorted(right.keys() - left.keys())
    raw_drift = []
    normalized_drift = []
    for item_key in sorted(left.keys() & right.keys()):
        if left[item_key]["raw_sha256"] != right[item_key]["raw_sha256"]:
            raw_drift.append(item_key)
        if left[item_key]["normalized_sha256"] != right[item_key]["normalized_sha256"]:
            normalized_drift.append(item_key)
    return {
        "ok": not missing and not extra and not normalized_drift,
        "missing": missing,
        "extra": extra,
        "raw_drift": raw_drift,
        "normalized_drift": normalized_drift,
    }


def cross_source_comparisons(records: Sequence[Mapping[str, object]], scenarios: Sequence[str]) -> list[dict]:
    """Compare working and reference pre-solver artifacts stage by stage."""
    stages = ("a1", "config", "a2", "otoole", "compiled")
    grouped: dict[tuple[str, str, str], dict[str, Mapping[str, object]]] = {}
    for record in records:
        source = str(record["source"])
        if source not in {"working_tracked", "reference_generated"}:
            continue
        key = (source, str(record["scenario"]), str(record["stage"]))
        grouped.setdefault(key, {})[str(record["path"])] = record
    comparisons = []
    for scenario in scenarios:
        for stage in stages:
            left_all = grouped.get(("working_tracked", scenario, stage), {})
            right_all = grouped.get(("reference_generated", scenario, stage), {})
            left = {path: record for path, record in left_all.items() if not excluded_from_exact_comparison(path)}
            right = {path: record for path, record in right_all.items() if not excluded_from_exact_comparison(path)}
            left_paths, right_paths = set(left), set(right)
            common = left_paths & right_paths
            raw_drift = sum(left[path]["raw_sha256"] != right[path]["raw_sha256"] for path in common)
            normalized_drift = sum(
                left[path]["normalized_sha256"] != right[path]["normalized_sha256"] for path in common
            )
            if not left:
                status = "missing-working"
            elif not right:
                status = "missing-reference"
            elif left_paths != right_paths:
                status = "file-set-drift"
            elif normalized_drift:
                status = "normalized-drift"
            elif raw_drift:
                status = "normalized-exact"
            else:
                status = "exact"
            comparisons.append({
                "scenario": scenario,
                "stage": stage,
                "status": status,
                "working_files": len(left),
                "reference_files": len(right),
                "common_files": len(common),
                "working_only": len(left_paths - right_paths),
                "reference_only": len(right_paths - left_paths),
                "working_excluded": len(left_all) - len(left),
                "reference_excluded": len(right_all) - len(right),
                "raw_drift": raw_drift,
                "normalized_drift": normalized_drift,
            })
    return comparisons


def excluded_from_exact_comparison(path: str) -> bool:
    name = Path(path).name.upper()
    return (
        "_PRE_" in name
        or "_PREPATCH_" in name
        or name.startswith("APPLY_PATCHES_CHANGES_")
    )


def _write_csv(path: Path, rows: Sequence[Mapping[str, object]], columns: Sequence[str]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def _coverage_rows(
    inventory: Sequence[Mapping[str, object]],
    working: Path,
    reference: Path | None,
    candidate: Path | None,
) -> list[dict]:
    work_tracked = tracked_files(working)
    rows = []
    for item in inventory:
        scenario = str(item["name"])
        row: dict[str, object] = {
            "scenario": scenario,
            "tier": item.get("tier", ""),
            "source_scenario": item.get("source_scenario", ""),
            "recipe": item.get("recipe", ""),
        }
        sources = (
            ("working_tracked", working, True, work_tracked),
            ("reference_generated", reference, False, None),
            ("candidate_generated", candidate, False, None),
        )
        for label, repo, tracked_only, tracked in sources:
            if repo is None:
                coverage = {"a1": False, "config": False, "a2": False, "otoole": False, "compiled_txt_count": 0, "output_csv_count": 0}
            else:
                coverage = scenario_coverage(repo, scenario, tracked_only=tracked_only, tracked=tracked)
            for field, value in coverage.items():
                row[f"{label}_{field}"] = value
        row["exact_static_status"] = (
            "comparable" if reference and row["working_tracked_a1"] and row["reference_generated_a1"]
            and row["working_tracked_config"] and row["reference_generated_config"] else "missing-evidence"
        )
        row["historical_output_status"] = "available" if int(row["reference_generated_output_csv_count"]) else "missing"
        row["solver_equivalence_status"] = "pending-cplex-baseline"
        rows.append(row)
    return rows


def capture_evidence(
    working: Path,
    out: Path,
    *,
    reference: Path | None = None,
    candidate: Path | None = None,
    scenario_file: Path = DEFAULT_SCENARIOS,
) -> dict:
    working = working.resolve()
    reference = reference.resolve() if reference else None
    candidate = candidate.resolve() if candidate else None
    inventory = load_scenarios(scenario_file)
    names = [str(item["name"]) for item in inventory]
    discovery = discover_scenarios(working, inventory)
    if not discovery_passes(discovery):
        raise RegressionError("working repository does not have the exact authoritative 20-scenario A1/config inventory")
    out.mkdir(parents=True, exist_ok=True)
    records = collect_artifact_records("working_tracked", working, names, tracked_only=True, include_generated=True)
    if reference:
        records.extend(collect_artifact_records("reference_generated", reference, names, tracked_only=False, include_generated=True))
    if candidate:
        records.extend(collect_artifact_records("candidate_generated", candidate, names, tracked_only=False, include_generated=True))
    records.sort(key=lambda item: (item["source"], item["scenario"], item["stage"], item["path"]))
    metrics = aggregate_records(records)
    coverage = _coverage_rows(inventory, working, reference, candidate)
    comparisons = cross_source_comparisons(records, names)
    manifest = {
        "schema_version": 1,
        "evidence_kind": "offline-static-and-historical-output",
        "normalizer_version": 1,
        "scenario_count": len(names),
        "scenarios": inventory,
        "working": git_metadata(working),
        "reference": git_metadata(reference) if reference else None,
        "candidate": git_metadata(candidate) if candidate else None,
        "protected_working_tree": protected_snapshot(working),
        "runtime": {"python": platform.python_version(), "platform": platform.system()},
        "solver_execution": "not-performed",
        "equivalence_scope": {
            "exact_pre_solver_static": "hashes and normalized hashes for artifacts that exist in both sources",
            "historical_outputs": "reference-only hashes; provenance is incomplete and coverage is not all-20",
            "full_solver_behavior": "pending a coherent all-20 CPLEX-backed baseline",
        },
        "comparison_summary": dict(sorted(Counter(item["status"] for item in comparisons).items())),
    }
    (out / "manifest.json").write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    hash_columns = ("source", "scenario", "stage", "path", "size", "raw_sha256", "normalized_sha256", "normalization")
    _write_csv(out / "hashes.csv", records, hash_columns)
    metric_columns = ("source", "scenario", "stage", "file_count", "total_bytes", "aggregate_normalized_sha256")
    _write_csv(out / "metrics.csv", metrics, metric_columns)
    comparison_columns = (
        "scenario", "stage", "status", "working_files", "reference_files", "common_files",
        "working_only", "reference_only", "working_excluded", "reference_excluded",
        "raw_drift", "normalized_drift",
    )
    _write_csv(out / "comparisons.csv", comparisons, comparison_columns)
    coverage_columns = tuple(coverage[0].keys())
    _write_csv(out / "coverage.csv", coverage, coverage_columns)
    return {
        "manifest": manifest,
        "hashes": records,
        "metrics": metrics,
        "coverage": coverage,
        "comparisons": comparisons,
    }


def _read_csv_dicts(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def cleanup_acceptance_report(
    inventory: Sequence[Mapping[str, object]],
    coverage: Sequence[Mapping[str, str]],
    comparisons: Sequence[Mapping[str, str]],
    manifest: Mapping[str, object] | None = None,
) -> dict:
    """Evaluate the user-approved 16-scenario offline cleanup gate."""
    accepted = scenarios_for_scope(inventory, CLEANUP_ACCEPTANCE_SCOPE)
    accepted_names = [str(item["name"]) for item in accepted]
    excluded = [
        {
            "name": str(item["name"]),
            "reason": str(item["cleanup_exclusion_reason"]),
        }
        for item in inventory
        if not item["cleanup_acceptance"]
    ]
    coverage_by_scenario = {str(row.get("scenario", "")): row for row in coverage}
    comparison_by_key = {
        (str(row.get("scenario", "")), str(row.get("stage", ""))): row
        for row in comparisons
    }
    failures: list[str] = []

    for item in inventory:
        scenario = str(item["name"])
        row = coverage_by_scenario.get(scenario)
        if row is None:
            failures.append(f"{scenario}: missing preservation coverage row")
            continue
        for field in (
            "working_tracked_a1",
            "working_tracked_config",
            "reference_generated_a1",
            "reference_generated_config",
        ):
            if str(row.get(field, "")).lower() != "true":
                failures.append(f"{scenario}: preservation field {field} is not present")

    status_counts: Counter[str] = Counter()
    for scenario in accepted_names:
        row = coverage_by_scenario.get(scenario)
        if row is None:
            failures.append(f"{scenario}: missing cleanup-acceptance coverage row")
            continue
        for stage in STATIC_ACCEPTANCE_STAGES:
            for source in ("working_tracked", "reference_generated"):
                field = f"{source}_{stage}"
                if str(row.get(field, "")).lower() != "true":
                    failures.append(f"{scenario}: required field {field} is not present")
            comparison = comparison_by_key.get((scenario, stage))
            if comparison is None:
                failures.append(f"{scenario}/{stage}: missing static comparison")
            else:
                status = str(comparison.get("status", ""))
                status_counts[status] += 1
                if status not in {"exact", "normalized-exact"}:
                    failures.append(f"{scenario}/{stage}: unacceptable comparison status {status}")
        for field in (
            "reference_generated_compiled_txt_count",
            "reference_generated_output_csv_count",
        ):
            try:
                present = int(str(row.get(field, "0"))) > 0
            except ValueError:
                present = False
            if not present:
                failures.append(f"{scenario}: required historical field {field} is absent")

    manifest = manifest or {}
    return {
        "schema_version": 1,
        "scope": CLEANUP_ACCEPTANCE_SCOPE,
        "ok": not failures,
        "preservation_scenario_count": len(inventory),
        "cleanup_acceptance_scenario_count": len(accepted_names),
        "cleanup_acceptance_scenarios": accepted_names,
        "excluded_scenarios": excluded,
        "required_static_stages": list(STATIC_ACCEPTANCE_STAGES),
        "static_comparison_summary": dict(sorted(status_counts.items())),
        "historical_reference_requirement": (
            "compiled text and direct output CSVs present for each accepted scenario"
        ),
        "solver_execution": "not-performed",
        "solver_equivalence": "pending a source-bound CPLEX baseline",
        "source_evidence": {
            "working": manifest.get("working"),
            "reference": manifest.get("reference"),
            "protected_working_tree": manifest.get("protected_working_tree"),
        },
        "failures": sorted(set(failures)),
    }


def evaluate_cleanup_acceptance(evidence_dir: Path, scenario_file: Path = DEFAULT_SCENARIOS) -> dict:
    required = ("manifest.json", "coverage.csv", "comparisons.csv")
    missing = required_files_report(evidence_dir, required)["missing"]
    if missing:
        raise RegressionError("acceptance evidence is missing: " + ", ".join(missing))
    manifest = json.loads((evidence_dir / "manifest.json").read_text(encoding="utf-8"))
    return cleanup_acceptance_report(
        load_scenarios(scenario_file),
        _read_csv_dicts(evidence_dir / "coverage.csv"),
        _read_csv_dicts(evidence_dir / "comparisons.csv"),
        manifest,
    )


def compare_evidence_dirs(baseline: Path, candidate: Path) -> dict:
    required = ("manifest.json", "hashes.csv", "metrics.csv", "coverage.csv", "comparisons.csv")
    left_missing = required_files_report(baseline, required)["missing"]
    right_missing = required_files_report(candidate, required)["missing"]
    if left_missing or right_missing:
        return {"ok": False, "baseline_missing": left_missing, "candidate_missing": right_missing}
    result = compare_hash_records(_read_csv_dicts(baseline / "hashes.csv"), _read_csv_dicts(candidate / "hashes.csv"))
    result["baseline_missing"] = []
    result["candidate_missing"] = []
    return result


def verify_protected(repo: Path, manifest_path: Path) -> dict:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    expected = manifest.get("protected_working_tree")
    actual = protected_snapshot(repo.resolve())
    return {"ok": expected == actual, "expected": expected, "actual": actual}


def _format_set(values: Iterable[str]) -> str:
    values = sorted(values)
    return ", ".join(values) if values else "none"


def command_discover(args: argparse.Namespace) -> int:
    inventory = load_scenarios(args.scenarios_file)
    preservation = discover_scenarios(args.repo, inventory)
    scoped = scenarios_for_scope(inventory, args.scope)
    result = discover_scenarios(args.repo, scoped)
    print(
        f"scope={args.scope} expected={len(result['expected'])} "
        f"a1={len(result['a1'] & result['expected'])} "
        f"configs={len(result['configs'] & result['expected'])}"
    )
    print(
        f"a2={len(result['a2'] & result['expected'])} "
        f"otoole={len(result['otoole'] & result['expected'])}"
    )
    for key in ("missing_a1", "missing_configs", "missing_a2", "missing_otoole"):
        print(f"{key}={_format_set(result[key])}")
    preservation_passed = discovery_passes(preservation)
    passed = (
        preservation_passed
        if args.scope in PRESERVATION_SCOPES
        else preservation_passed and cleanup_acceptance_discovery_passes(result)
    )
    if passed and args.scope in PRESERVATION_SCOPES:
        print("PASS exact authoritative 20-scenario preservation discovery")
    elif passed:
        print("PASS 16-scenario cleanup acceptance discovery with 20-scenario preservation")
    else:
        print("FAIL scenario discovery")
    return 0 if passed else 1


def command_capture(args: argparse.Namespace) -> int:
    evidence = capture_evidence(
        args.repo,
        args.out,
        reference=args.reference_repo,
        candidate=args.candidate_repo,
        scenario_file=args.scenarios_file,
    )
    counts = Counter((item["source"], item["stage"]) for item in evidence["hashes"])
    print(f"captured {len(evidence['hashes'])} compact file hashes in {args.out}")
    for (source, stage), count in sorted(counts.items()):
        print(f"{source}/{stage}: {count}")
    return 0


def command_compare(args: argparse.Namespace) -> int:
    result = compare_evidence_dirs(args.baseline, args.candidate)
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0 if result.get("ok") else 1


def command_gate(args: argparse.Namespace) -> int:
    result = evaluate_cleanup_acceptance(args.evidence, args.scenarios_file)
    payload = json.dumps(result, indent=2, sort_keys=True) + "\n"
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(payload, encoding="utf-8")
    print(payload, end="")
    return 0 if result["ok"] else 1


def command_verify_protected(args: argparse.Namespace) -> int:
    result = verify_protected(args.repo, args.manifest)
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0 if result["ok"] else 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    discover = sub.add_parser("discover", help="verify exact A1/config scenario discovery")
    discover.add_argument("--repo", type=Path, default=Path("."))
    discover.add_argument(
        "--scope",
        choices=["regression", "preservation", CLEANUP_ACCEPTANCE_SCOPE],
        default="regression",
    )
    discover.add_argument("--scenarios-file", type=Path, default=DEFAULT_SCENARIOS)
    discover.set_defaults(func=command_discover)

    capture = sub.add_parser("capture", help="capture compact offline evidence")
    capture.add_argument("--repo", type=Path, default=Path("."))
    capture.add_argument("--reference-repo", type=Path)
    capture.add_argument("--candidate-repo", type=Path)
    capture.add_argument("--out", type=Path, required=True)
    capture.add_argument("--scenarios-file", type=Path, default=DEFAULT_SCENARIOS)
    capture.set_defaults(func=command_capture)

    compare = sub.add_parser("compare", help="compare two compact evidence directories")
    compare.add_argument("--baseline", type=Path, required=True)
    compare.add_argument("--candidate", type=Path, required=True)
    compare.add_argument("--profile", choices=["strict"], default="strict")
    compare.set_defaults(func=command_compare)

    gate = sub.add_parser("gate", help="evaluate an explicit offline cleanup acceptance scope")
    gate.add_argument("--evidence", type=Path, required=True)
    gate.add_argument("--scope", choices=[CLEANUP_ACCEPTANCE_SCOPE], default=CLEANUP_ACCEPTANCE_SCOPE)
    gate.add_argument("--scenarios-file", type=Path, default=DEFAULT_SCENARIOS)
    gate.add_argument("--out", type=Path)
    gate.set_defaults(func=command_gate)

    protected = sub.add_parser("verify-protected", help="verify protected trees against a manifest")
    protected.add_argument("--repo", type=Path, default=Path("."))
    protected.add_argument("--manifest", type=Path, required=True)
    protected.set_defaults(func=command_verify_protected)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args))
    except RegressionError as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

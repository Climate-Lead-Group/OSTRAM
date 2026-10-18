"""Regenerate the generated block of the scenario catalogue page. Run from the repository root.

The block documents the delivered scenarios: the registered scenarios that have a
row in the origin table of the page, in registry order. A scenario registered later
appears only after its row is added to the table.
"""
from pathlib import Path
import re
import sys
import yaml

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from ostram.paths import resolve_paths  # noqa: E402

PAGE = ROOT / "docs" / "reference" / "scenarios" / "scenario_catalogue.md"
START, END = "<!-- generated: start -->", "<!-- generated: end -->"


def origin_rows(page):
    """Return the scenario names that have a row in the origin table of ``page``."""
    origin = page.split("## Origin of each scenario", 1)[1]
    return re.findall(r"^\| `([A-Za-z0-9_]+)` \|", origin, flags=re.MULTILINE)


def _edits(edits):
    groups = {}
    for e in edits:
        tech = e.get("tech") or (", ".join(e["techs"]) if "techs" in e else e["tech_prefix"] + "*")
        groups.setdefault((e["sheet"], e["param"], "values" if "values" in e else e.get("op", "")), []).append(tech)
    rows = [f"| {s} | `{p}` | {op} | {', '.join(t)} |" for (s, p, op), t in groups.items()]
    return ["", "| Sheet | Parameter | Operation | Technologies |", "|---|---|---|---|"] + rows


def generated_block(names=None):
    """Return the generated block for ``names``, or for the delivered scenarios."""
    paths = resolve_paths(project_root=ROOT)
    setup = yaml.safe_load(paths.scenario_registry.read_text(encoding="utf-8"))
    execution = yaml.safe_load((ROOT / "config/execution/Config_MOMF_T1_AB.yaml").read_text(encoding="utf-8"))
    registry = setup["scenario_registry"]
    roots = {r["name"]: r for r in registry["root_scenarios"]}
    registered = list(registry["support_scenarios"]) + list(registry["decision_scenarios"])
    if names is None:
        listed = set(origin_rows(PAGE.read_text(encoding="utf-8")))
        names = [name for name in registered if name in listed]
    out = [START]
    for name in names:
        scenario_file = paths.scenario_config_root / name / "scenario.yaml"
        block = {} if name in roots else yaml.safe_load(scenario_file.read_text(encoding="utf-8"))
        root = name if name in roots else block["base_scenario"]
        reused = block.get("reuse_edits_from")
        inherited = (yaml.safe_load((paths.scenario_config_root / reused / "scenario.yaml").read_text(encoding="utf-8"))
                     .get("edits", []) if reused else [])
        seeds = [d["scenario"] for d in roots[root]["dependencies"] if d.get("type") == "result"]
        policy = []
        if root in setup["policies"].get("historical_trade_roots", []):
            policy.append("the historical trade constraints are active")
        if name in execution.get("cplex_numerical_emphasis_scenarios", []):
            policy.append("CPLEX numerical emphasis")
        if cap := block.get("gross_import_cap"):
            policy.append(f"gross imports capped at {cap['fraction']} of annual demand in {cap['first_year']} to "
                          f"{cap['last_year']} for {', '.join(cap['countries'])}")
        out += [f"### {name}", ""]
        if name in roots:
            out.append(f"- Parent: none. `{name}` is a {roots[name]['role']} root.")
        else:
            overlay = block.get("direction_overlay")
            out += [f"- Parent: `{root}`.",
                    f"- Shared VRE ceilings: {'applied' if block.get('apply_vre_ceiling_layer', True) else 'skipped'}.",
                    f"- Scenario edits: {len(block.get('edits', []))}"
                    + (f", after the {len(inherited)} edits of `{reused}`." if reused else "."),
                    "- Overlay: " + (f"`config/scenarios/{name}/{overlay}` from {block['direction_study_start_year']}."
                                     if overlay else "none.")]
        out += [f"- B2 policy: {'; '.join(policy) or 'none'}.",
                f"- Seed needs: {', '.join('a solved `' + s + '` result' for s in seeds) or 'none'}."]
        if name not in roots:
            out.append(f"- Description: {block.get('description') or block.get('rule')}")
            out += _edits(inherited + block.get("edits", [])) if inherited or block.get("edits") else []
        out.append("")
    return "\n".join(out + [END])


def current_block(text):
    return text[text.index(START):text.index(END) + len(END)]


if __name__ == "__main__":
    text = PAGE.read_text(encoding="utf-8")
    PAGE.write_text(text.replace(current_block(text), generated_block()), encoding="utf-8", newline="\n")
    print(f"regenerated {PAGE.relative_to(ROOT)}")

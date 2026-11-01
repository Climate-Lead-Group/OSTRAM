"""Generated parts of the site pages: tables, the worked example, the stepper and diagrams.

Each block is a function ``block_<name>(site, page, argument)`` that returns HTML.
Every number a block prints is a value from :mod:`ostram.site.values`, printed
with :func:`value_span`, so the page names its source and the tests can check it.
The source entries come from :mod:`ostram.site.sources`; :func:`page_sources`
lists the ones a page used.
"""

from __future__ import annotations

import html
import importlib
import posixpath
import re

from ostram.site import sources
from ostram.site.values import INTERNAL_CATEGORIES, literal_assignment, rule_modules, code_counts

e = html.escape

# What each internal category carries, as docs/guides/internal-expansion.md describes it.
CATEGORY_TEXT = {
    "TRANSMISSION_PLANNED_RE": ("renewable side, existing lines", "00"),
    "TRANSMISSION_REPOT_RE": ("renewable side, repowered lines", "00"),
    "TRANSMISSION_UNPLANNED_RE": ("renewable side, new lines", "00"),
    "TRANSMISSION_PLANNED_NONRENO": ("other plants and imports, existing lines", "01"),
    "TRANSMISSION_REPOT_NR": ("other plants and imports, repowered lines", "01"),
    "TRANSMISSION_UNPLANNED_NR": ("other plants and imports, new lines", "01"),
}

# Chain names in the "Order of the rules" table of docs/rules.md and the module of each.
CHAIN_MODULES = {
    "retirement": "set_retirement_schedule.py",
    "build cap": "build_cap_rule.py",
    "minimum capacity floors": "set_min_capacity_floors.py",
    "relax interconnectors": "relax_interconnectors.py",
    "VRE targets": "set_vre_targets.py",
    "storage minimum investment": "add_storage_min_investment.py",
}
SHARED_MODULES = (
    "apply_interconnector_costs.py",
    "apply_internal_transmission.py",
    "apply_internal_tx_losses.py",
)


def format_value(value, kind: str = "") -> str:
    """Return ``value`` as the site prints it."""

    if kind == "pct":
        return f"{float(value) * 100:.10g}%"
    if kind == "int":
        return f"{int(value):,}"
    if kind == "year":
        return str(int(value))
    if kind.startswith("f") and kind[1:].isdigit():
        return f"{float(value):,.{int(kind[1:])}f}"
    if isinstance(value, bool):
        return "yes" if value else "no"
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else f"{value:.10g}"
    if isinstance(value, int):
        return str(value) if 1000 <= value <= 2100 else f"{value:,}"
    return str(value)


def value_span(site, key: str, kind: str = "") -> str:
    """Return one value wrapped so that readers and tests can see its source and file.

    A value left out for want of a record prints nothing.
    """

    if key in site.left_out:
        return ""
    item = site.value(key)
    text = e(format_value(item.value, kind))
    title = e(
        sources.one_line_entry(site, key) + "\n"
        + f"Read from {item.source}" + (f" ({item.note})" if item.note else "")
    )
    kind_attr = f' data-f="{kind}"' if kind else ""
    numbers = " ".join(str(n) for n in refs(site, key))
    return f'<span class="v" data-v="{e(key)}"{kind_attr} data-s="{numbers}" title="{title}">{text}</span>'


def refs(site, key: str) -> list[int]:
    """Return the numbers of a value's sources in the list at the end of the page, adding new ones."""

    numbers = []
    for unit in sources.units(site, key):
        if unit not in site.page_units:
            site.page_units.append(unit)
        numbers.append(site.page_units.index(unit) + 1)
    return numbers


def _ref_links(numbers: list[int]) -> str:
    return " ".join(f'<a href="#src-{n}">[{n}]</a>' for n in numbers)


def _linkify_entry(text: str) -> str:
    """Return record text with each URL a link, ending each URL as :func:`sources.urls_of` does."""

    parts, last = [], 0
    for url in sources.urls_of(text):
        start = text.find(url, last)
        if start < 0:
            continue
        parts.append(e(text[last:start]))
        parts.append(f'<a href="{e(url)}" rel="noopener">{e(url)}</a>')
        last = start + len(url)
    parts.append(e(text[last:]))
    return "".join(parts)


def _linkify_urls(text: str, urls) -> str:
    out = _linkify_entry(text)
    missing = [url for url in urls if url not in text]
    return out + "".join(f' <a href="{e(url)}" rel="noopener">{e(url)}</a>' for url in missing)


def _unit_html(unit: tuple) -> str:
    if unit[0] == "citation":
        return _linkify_urls(unit[1], unit[2])
    if unit[0] == "assumption":
        return f"<em>assumption</em> {_linkify_entry(unit[1])}"
    _, label, path, url = unit
    return f'{e(label)}: <a href="{e(url)}" rel="noopener"><code>{e(path)}</code></a>'


def entry_html(site, key: str) -> str:
    """Return the entry of a value for a table cell, with links to the list of sources."""

    entry = site.entries[key]
    numbers = refs(site, key)
    if entry["kind"] == "setting":
        return f'<a href="{e(entry["url"])}" rel="noopener">{e(entry["label"])}</a> {_ref_links(numbers)}'
    if entry["kind"] == "computed":
        inputs = ", ".join(_code(name) for name in entry["inputs"])
        return f"computed: {e(entry['formula'])}, from {inputs} {_ref_links(numbers)}"
    if entry["kind"] == "assumption":
        sentences = " ".join(_linkify_entry(sentence) for sentence in entry["assumptions"])
        return f"<em>assumption</em> {sentences} {_ref_links(numbers)}"
    return f"source {_ref_links(numbers)}"


def page_sources(site) -> str:
    """Return the section "Sources on this page" for the sources the page's values used."""

    if not site.page_units:
        return ""
    items = "".join(f'<li id="src-{n}">{_unit_html(unit)}</li>' for n, unit in enumerate(site.page_units, 1))
    return (
        '\n<h2 id="sources-on-this-page">Sources on this page</h2>\n'
        "<p>Each value on this page names its sources in its hover text. External sources link to the "
        "document; an <em>assumption</em> is quoted as the repository records it; a model definition, "
        "scenario setting or repository count links to the file it is read from.</p>"
        f'<ol class="page-sources">{items}</ol>'
    )


def rel(from_page: str, to_page: str) -> str:
    start = posixpath.dirname(from_page) or "."
    return posixpath.relpath(to_page, start)


def applications(site) -> list[str]:
    """Return the registered setups the site presents as applications, in site order."""

    return [app["setup"] for app in site.config["applications"] if app["setup"] in site.setups]


def app_config(site, setup_id: str) -> dict:
    for app in site.config["applications"]:
        if app["setup"] == setup_id:
            return app
    return {"setup": setup_id, "name": site.setups[setup_id].title}


def has(site, key: str) -> bool:
    setup_id = key.split(".", 1)[0]
    return setup_id in site.setups and key in site.setups[setup_id].values


def _table(header: list[str], rows: list[list[str]], css: str = "") -> str:
    head = "".join(f"<th>{cell}</th>" for cell in header)
    body = "".join("<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>" for row in rows)
    return f'<div class="table-wrap"><table class="{css}"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


def _code(text: str) -> str:
    return f"<code>{e(text)}</code>"


def _source_link(site, path: str) -> str:
    return f'<a href="{e(site.blob(path))}" rel="noopener"><code>{e(path)}</code></a>'


def _years(site, setup_id: str) -> str:
    return f"{value_span(site, f'{setup_id}.first_year')} to {value_span(site, f'{setup_id}.last_year')}"


# Home and applications -------------------------------------------------------


def block_app_cards(site, page, argument):
    cards = []
    for app in site.config["applications"]:
        setup_id = app["setup"]
        link = rel(page.path, f"applications/{setup_id}.html")
        if setup_id not in site.setups:
            cards.append(
                f'<article class="card app pending"><p class="eyebrow">{e(app["route"])}</p>'
                f'<h3><a href="{link}">{e(app["name"])}</a></h3>'
                f'<p>{e(app["placeholder"])}</p><p class="status">Not yet registered: the counts appear '
                "when the setup is added to the setup registry.</p></article>"
            )
            continue
        stats = [
            (value_span(site, f"{setup_id}.nodes"), "nodes"),
            (value_span(site, f"{setup_id}.interconnectors"), "interconnector technologies"),
            (value_span(site, f"{setup_id}.scenarios"), "scenarios"),
            (f'{value_span(site, f"{setup_id}.first_year")}&#8211;{value_span(site, f"{setup_id}.last_year")}', "years"),
            (value_span(site, f"{setup_id}.timeslices"), "time slices"),
        ]
        tiles = "".join(f'<div class="stat"><span class="num">{num}</span><span class="lbl">{label}</span></div>' for num, label in stats)
        cards.append(
            f'<article class="card app"><p class="eyebrow">{e(app["route"])}</p>'
            f'<h3><a href="{link}">{e(app["name"])}</a></h3><p>{e(app["summary"])}</p>'
            f'<div class="stats">{tiles}</div></article>'
        )
    return f'<div class="apps">{"".join(cards)}</div>'


def block_setup_list(site, page, argument):
    rows = []
    for setup_id in site.setups:
        app = app_config(site, setup_id)
        rows.append(
            [
                f'<a href="{rel(page.path, f"applications/{setup_id}.html")}">{e(app["name"])}</a>',
                _code(setup_id),
                value_span(site, f"{setup_id}.nodes") if has(site, f"{setup_id}.nodes") else "",
                value_span(site, f"{setup_id}.interconnectors") if has(site, f"{setup_id}.interconnectors") else "",
                value_span(site, f"{setup_id}.scenarios") if has(site, f"{setup_id}.scenarios") else "",
                _source_link(site, site.setups[setup_id].manifest),
            ]
        )
    for app in site.config["applications"]:
        if app["setup"] not in site.setups:
            rows.append(
                [
                    f'<a href="{rel(page.path, "applications/" + app["setup"] + ".html")}">{e(app["name"])}</a>',
                    _code(app["setup"]),
                    "", "", "", "not registered yet",
                ]
            )
    return _table(["Application", "Setup id", "Nodes", "Interconnectors", "Scenarios", "Manifest"], rows)


def block_setup_summary(site, page, argument):
    setup_id = page.setup
    info = site.setups[setup_id]
    stats = [
        ("Nodes", value_span(site, f"{setup_id}.nodes")),
        ("Interconnector technologies", value_span(site, f"{setup_id}.interconnectors")),
        ("Internal grid technologies", value_span(site, f"{setup_id}.internal_grid_technologies")),
        ("Years", f"{value_span(site, f'{setup_id}.years')} ({_years(site, setup_id)})"),
        (
            "Time slices",
            f"{value_span(site, f'{setup_id}.timeslices')} ({value_span(site, f'{setup_id}.seasons')} seasons, "
            f"{value_span(site, f'{setup_id}.day_types')} day type, "
            f"{value_span(site, f'{setup_id}.daily_time_brackets')} daily time brackets)",
        ),
        (
            "Scenarios",
            f"{value_span(site, f'{setup_id}.scenarios')}: {value_span(site, f'{setup_id}.support_scenarios')} support and "
            f"{value_span(site, f'{setup_id}.decision_scenarios')} decision. "
            f"{value_span(site, f'{setup_id}.root_scenarios')} are roots and "
            f"{value_span(site, f'{setup_id}.derived_scenarios')} are derived from a root",
        ),
    ]
    rows = [[label, value] for label, value in stats]
    nodes = [[_code(node), e(info.node_labels.get(node, ""))] for node in info.nodes]
    links = []
    for tech in info.interconnectors:
        a, b = tech[3:8], tech[8:13]
        links.append([_code(tech), e(info.node_labels.get(a, a)), e(info.node_labels.get(b, b))])
    authorities = [[_code(role), _source_link(site, path) if not path.endswith("/") else _code(path)] for role, path in _authority_rows(site, setup_id)]
    return (
        _table(["What", "Count"], rows, "kv")
        + '<h3 id="nodes">Nodes</h3>'
        + _table(["Node", "Name"], nodes)
        + '<h3 id="interconnectors">Interconnector technologies</h3>'
        + "<p>Mode 1 carries power from the first node in the name to the second; mode 2 carries it back.</p>"
        + _table(["Technology", "First node", "Second node"], links)
        + '<h3 id="inputs">Input files</h3>'
        + f"<p>The setup manifest {_source_link(site, info.manifest)} names every input of this application.</p>"
        + _table(["Role", "File or folder"], authorities)
    )


def _authority_rows(site, setup_id: str) -> list[tuple[str, str]]:
    from ostram.site.values import relative
    from ostram.setups import load_manifest

    info = site.setups[setup_id]
    manifest = load_manifest(site.paths.project_root / info.manifest, expected_setup=setup_id)
    rows = []
    for role, path in manifest.source_paths(site.paths, require_exists=False).items():
        spec = manifest.authorities[role]
        shown = relative(path, site.paths.project_root)
        if spec.namespace == "package":
            shown = "ostram/" + shown if not shown.startswith("ostram/") else shown
        rows.append((role, shown + ("/" if spec.kind == "directory" else "")))
    return rows


def block_setup_title(site, page, argument):
    return e(app_config(site, page.setup)["name"]) if page.setup in site.setups or any(
        a["setup"] == page.setup for a in site.config["applications"]) else e(page.setup)


def block_setup_intro(site, page, argument):
    app = app_config(site, page.setup) if page.setup in site.setups else next(
        a for a in site.config["applications"] if a["setup"] == page.setup)
    if "route" in app:
        return f'<p class="eyebrow">{e(app["route"])}</p><p class="lead">{e(app["summary"])}</p>'
    info = site.setups[page.setup]
    from ostram.setups import load_manifest

    manifest = load_manifest(site.paths.project_root / info.manifest, expected_setup=page.setup)
    text = str(manifest.metadata.get("description") or info.title)
    return f'<p class="eyebrow">Registered setup</p><p class="lead">{e(" ".join(text.split()))}</p>'


def block_lac_status(site, page, argument):
    argument = page.setup if argument == "@" else argument
    app = next(a for a in site.config["applications"] if a["setup"] == argument)
    if argument in site.setups:
        return f'<p class="note">The {e(app["name"])} setup is registered; its values below are read from its files.</p>'
    return (
        f'<div class="placeholder"><p><strong>{e(app["name"])}: coming.</strong> {e(app["placeholder"])}</p>'
        f'<p>This part fills itself when a setup with the id {_code(argument)} is added to '
        f'{_source_link(site, "config/setups/registry.json")}: the same templates that build the South Asia '
        "pages then read its files.</p></div>"
    )


# The model as choices and rules ---------------------------------------------


def _node_groups(info, node: str) -> list[tuple[str, str, list[str]]]:
    techs = info.technologies_by_node.get(node, [])
    groups = [
        ("primary", "Primary supply", [t for t in techs if t.startswith(("MIN", "RNW")) and not t.startswith(("RNWTRN", "RNWRPO", "RNWNLI"))]),
        ("gen", "Power plants", [t for t in techs if t.startswith("PWR") and not t.startswith("PWRTRN")]),
        ("storage", "Storage", [t for t in techs if t.startswith(("LDS", "SDS"))]),
        ("grid", "Internal grid", [t for t in techs if t[:6] in ("RNWTRN", "RNWRPO", "RNWNLI", "PWRTRN", "TRNRPO", "TRNNLI")]),
        ("dispatch", "Dispatch", [t for t in techs if t.startswith("DSPTRN")]),
        ("link", "Interconnectors", [t for t in info.interconnectors if node in (t[3:8], t[8:13])]),
    ]
    return [g for g in groups if g[2]]


def block_choices_sheet(site, page, argument):
    example = site.config["worked_example"]
    setup_id, node = example["setup"], example["node"]
    info = site.setups[setup_id]
    label = info.node_labels.get(node, node)
    year = value_span(site, f"{setup_id}.last_year")
    groups = _node_groups(info, node)
    slices = info.timeslices
    head = '<th class="rowhead">Technology</th><th class="col-build">New capacity</th>' + "".join(
        f'<th class="col-slice">{e(ts)}</th>' for ts in slices
    )
    body = []
    for key, title, techs in groups:
        body.append(f'<tr class="group"><th colspan="{len(slices) + 2}">{e(title)}</th></tr>')
        for tech in techs:
            cells = f'<td class="c build" data-g="{key}" data-t="{e(tech)}"></td>' + "".join(
                f'<td class="c run" data-g="{key}" data-t="{e(tech)}"></td>' for _ in slices
            )
            body.append(f'<tr><th class="rowhead"><code>{e(tech)}</code></th>{cells}</tr>')
    ceiling_rule = ""
    ceilings = [t for t in info.vre_ceilings if t.endswith(node)]
    if ceilings:
        parts = ", ".join(f"{_code(t)} at most {value_span(site, f'{setup_id}.vre_ceiling.{t}')} GW" for t in ceilings)
        ceiling_rule = (
            f'<li><button type="button" data-rule="ceiling" data-hit="build:{" ".join(ceilings)}">VRE ceiling</button>'
            f"<span>The total capacity of solar and onshore wind stays below the shared ceiling of the setup: {parts}.</span></li>"
        )
    links = [t for t in info.interconnectors if node in (t[3:8], t[8:13])]
    residuals = ", ".join(
        f"{_code(t)} {value_span(site, f'{setup_id}.corridor.{t}.residual')} GW"
        for t in links if has(site, f"{setup_id}.corridor.{t}.residual")
    )
    rules = [
        ("balance", "run:dispatch link grid", "Energy balance",
         f"In each of the {value_span(site, f'{setup_id}.timeslices')} time slices, the power that reaches the demand fuel of {e(label)} equals its demand plus what it sends abroad."),
        ("capacity", "run:primary gen storage grid dispatch link", "Capacity limit",
         "In each time slice, a technology runs at most at its capacity times its capacity factor times the time slice's share of the year."),
        ("loss", "run:grid", "Internal grid loss",
         f"What leaves the internal grid is {value_span(site, f'{setup_id}.internal.efficiency')} of what enters it: the loss setting is {value_span(site, f'{setup_id}.internal.transmission_loss', 'pct')}."),
        ("cap", "build:gen", "Build cap", "New capacity of a plant in a year stays at or below its build cap, which the A3 rules write."),
        ("link", "run:link", "Link capacity",
         "Flow on each interconnector stays within its capacity. Existing capacity in the first year: " + (residuals or "none listed") + "."),
        ("reserve", "build:gen storage", "Reserve margin", "Firm capacity across the system covers peak demand with a margin. The model holds one reserve pool for every node together."),
        ("cost", "build:primary gen storage grid link run:primary gen storage grid dispatch link", "Least cost",
         f"Among all plans that keep every rule, the solver picks the one with the lowest discounted cost over {value_span(site, f'{setup_id}.years')} years."),
    ]
    items = "".join(
        f'<li><button type="button" data-rule="{key}" data-hit="{hit}">{title}</button><span>{text}</span></li>'
        for key, hit, title, text in rules
    ) + ceiling_rule
    return (
        f'<div class="sheet-wrap" data-sheet>'
        f'<p class="sheet-caption">Choices for <strong>{e(label)}</strong> ({_code(node)}) in {year}: each empty cell '
        f"is one number the solver chooses. Pick a rule to see which cells it ties together.</p>"
        f'<ol class="rules-list">{items}</ol>'
        f'<div class="table-wrap sheet"><table><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'
        f'<p class="small">Rows are the technologies of {_code(node)} in {_source_link(site, info.sources["taxonomy"])}; '
        f"columns are the time slices of {_source_link(site, info.sources['compilation_config'])}. Schematic: the cells hold no model output.</p></div>"
    )


def block_timeslice_grid(site, page, argument):
    setup_id = argument or site.config["worked_example"]["setup"]
    info = site.setups[setup_id]
    seasons, brackets = info.seasons, info.brackets
    if not seasons or not brackets or seasons * brackets != len(info.timeslices):
        return _table(["Time slice"], [[_code(ts)] for ts in info.timeslices])
    head = ["Season"] + [f"Part of day {n}" for n in range(1, brackets + 1)]
    rows = []
    for s in range(seasons):
        rows.append([f"Season {s + 1}"] + [_code(ts) for ts in info.timeslices[s * brackets:(s + 1) * brackets]])
    return _table(head, rows, "grid") + (
        f'<p class="small">The time slices of {_source_link(site, info.sources["compilation_config"])}, '
        "in the order the compiler lists them.</p>"
    )


# The rules behind a result -----------------------------------------------------


def _rules_table(site) -> tuple[dict[str, tuple[str, str]], dict[str, list[str]]]:
    text = (site.docs_root / "rules.md").read_text(encoding="utf-8")
    modules: dict[str, tuple[str, str]] = {}
    for line in text.split("## The eleven A3 rules", 1)[1].split("\n## ", 1)[0].splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) == 3 and cells[0].startswith("`"):
            modules[cells[0].strip("`")] = (cells[1], cells[2])
    chains: dict[str, list[str]] = {}
    for line in text.split("## Order of the rules", 1)[1].split("\n## ", 1)[0].splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) == 2 and cells[0].startswith("`"):
            chains[cells[0].strip("`")] = [part.strip() for part in cells[1].split(",")]
    return modules, chains


def _inline(site, page, markdown: str) -> str:
    from ostram.site.markdown import MarkdownRenderer

    def resolve(href: str) -> str:
        anchor = href.partition("#")[2]
        doc = "rules.md" if href.startswith("#") else posixpath.normpath(href.partition("#")[0])
        target = site.anchors.get((doc, anchor)) if anchor else site.whole_docs.get(doc)
        if target:
            return (rel(page.path, target) if target != page.path else "") + (f"#{anchor}" if anchor else "")
        return site.config["docs_url"] + doc[:-3] + ".html" + (f"#{anchor}" if anchor else "")

    rendered = MarkdownRenderer(resolve=resolve, doc_title=lambda t: t).render(markdown).html
    return re.sub(r"^<p>|</p>\s*$", "", rendered.strip())


def block_rules_stepper(site, page, argument):
    modules, chains = _rules_table(site)
    pin_roots = sorted(literal_assignment(
        site.paths.package_root / "pipeline" / "scenarios" / "orchestrator.py", "PWR_MIN_PIN_ROOT_SCENARIOS"))
    folder = "ostram/pipeline/scenarios/rules/"
    tabs, panels = [], []
    for index, (root, chain) in enumerate(chains.items()):
        steps = []
        for name in chain:
            module = CHAIN_MODULES.get(name)
            steps.append((f"Chain: {name}", module))
        for module in SHARED_MODULES:
            steps.append(("Shared rule", module))
        if root in pin_roots:
            steps.append(("Base-year pin", "apply_base_year_pin.py"))
        steps.append(("Internal transmission residuals", None))
        if root == "B_Optimised_VRE":
            steps.append(("B as a relaxation of A", "relax_to_reference.py"))
        cards = []
        for number, (stage, module) in enumerate(steps, 1):
            if module and module in modules:
                settings, what = modules[module]
                body = f"<p>{_inline(site, page, what)}</p><p class=\"small\">Settings: {_inline(site, page, settings)}</p>"
            elif stage == "Internal transmission residuals":
                body = (f'<p>{_inline(site, page, "Computes the base-year capacity of the internal grid of each node from the base-year inputs the model compiles; see [Internal transmission residuals](#internal-transmission-residuals).")}</p>')
            elif module == "relax_to_reference.py":
                body = f'<p>{_inline(site, page, "Takes the looser of its own bounds and those of `A_Calibrated_BAU` built in the same run; see [B as a relaxation of A](#b-as-a-relaxation-of-a).")}</p>'
            else:
                body = ""
            title = f"<code>{e(module)}</code>" if module else e(stage)
            cards.append(
                f'<li class="step"><span class="step-n">{number}</span><div><p class="eyebrow">{e(stage)}</p>'
                f"<h4>{title}</h4>{body}</div></li>"
            )
        selected = "true" if index == 0 else "false"
        tabs.append(f'<button type="button" role="tab" aria-selected="{selected}" data-tab="{e(root)}">{e(root)}</button>')
        hidden = "" if index == 0 else " hidden"
        panels.append(
            f'<div class="stepper-panel" data-panel="{e(root)}"{hidden}><ol class="steps">{"".join(cards)}</ol>'
            '<div class="stepper-nav"><button type="button" data-step="-1">Previous</button>'
            '<span class="stepper-count"></span><button type="button" data-step="1">Next</button></div></div>'
        )
    present = {name for name, _ in rule_modules(site.paths)}
    listed = sorted(set(modules) | set(CHAIN_MODULES.values()) | {"relax_to_reference.py"})
    missing = [m for m in listed if m not in present]
    if missing:
        raise ValueError(f"docs/rules.md names rule modules that do not exist: {missing}")
    return (
        f'<div class="stepper" data-stepper><div class="tabs" role="tablist">{"".join(tabs)}</div>{"".join(panels)}</div>'
        f'<p class="small">Steps read from the tables of {_source_link(site, "docs/rules.md")}; the base-year pin roots from '
        f"{_code('PWR_MIN_PIN_ROOT_SCENARIOS')} in {_source_link(site, 'ostram/pipeline/scenarios/orchestrator.py')}. "
        f"Each module lives in {_code(folder)}.</p>"
    )


def block_rule_modules(site, page, argument):
    modules, _ = _rules_table(site)
    rows = []
    for name, _doc in rule_modules(site.paths):
        described = "listed in the rules table" if name in modules else (
            "B as a relaxation of A" if name == "relax_to_reference.py" else "helper, not a rule")
        rows.append([_source_link(site, f"ostram/pipeline/scenarios/rules/{name}"), described])
    return _table(["Module", "Role"], rows)


# Inside a node -------------------------------------------------------------------


def block_node_chain(site, page, argument):
    setup_id = argument or site.config["worked_example"]["setup"]
    eff = value_span(site, f"{setup_id}.internal.efficiency")
    loss = value_span(site, f"{setup_id}.internal.transmission_loss", "pct")
    notes = {
        "supply": "Primary supply: mining (<code>MIN</code>) and renewable resource (<code>RNW</code>) technologies deliver fuel or resource to the plants of the node.",
        "rplants": "Renewable plants and storage discharge put their output on <code>ELC&lt;node&gt;00</code>. Storage also charges from this fuel.",
        "oplants": "Other plants and the backstop put their output on <code>ELC&lt;node&gt;01</code>. The backstop is a very expensive plant used only when nothing else can meet demand.",
        "e00": "<code>ELC&lt;node&gt;00</code>: renewable output, before the internal grid.",
        "e01": "<code>ELC&lt;node&gt;01</code>: other plant output and received imports, before the internal grid.",
        "rgrid": f"Renewable side of the internal grid: <code>RNWTRN</code>, <code>RNWRPO</code> and <code>RNWNLI</code> (existing, repowered and new lines). Output is {eff} of input.",
        "ogrid": f"Other side of the internal grid: <code>PWRTRN</code>, <code>TRNRPO</code> and <code>TRNNLI</code>. Output is {eff} of input, so imports pay this loss too.",
        "e02": "<code>ELC&lt;node&gt;02</code>: power that has crossed the internal grid.",
        "dsp1": "<code>DSPTRN</code> mode 1 moves power from <code>ELC&lt;node&gt;02</code> to the demand fuel. It has no cost and never limits a flow.",
        "e03": "<code>ELC&lt;node&gt;03</code>: the demand fuel. Demand is met here and exports leave from here.",
        "demand": f"Demand measured at the bus bar is written as its published value times one minus the loss ({loss}), because the published figure already includes grid losses.",
        "export": "Interconnector out: an export consumes <code>ELC&lt;node&gt;03</code> of the sending node.",
        "import": "Interconnector in: a received import arrives on <code>ELC&lt;node&gt;04</code> after the link's own loss.",
        "e04": "<code>ELC&lt;node&gt;04</code>: imports as received.",
        "dsp2": "<code>DSPTRN</code> mode 2 moves received imports to <code>ELC&lt;node&gt;01</code>, so they cross the internal grid before they meet demand.",
    }
    boxes = [
        ("supply", 10, 30, 120, "Primary supply", "MIN, RNW"),
        ("rplants", 160, 10, 140, "Renewable plants", "and storage"),
        ("oplants", 160, 110, 140, "Other plants", "and backstop"),
        ("e00", 330, 10, 90, "ELC..00", ""),
        ("e01", 330, 110, 90, "ELC..01", ""),
        ("rgrid", 450, 10, 150, "Internal grid", "RNWTRN, RPO, NLI"),
        ("ogrid", 450, 110, 150, "Internal grid", "PWRTRN, RPO, NLI"),
        ("e02", 630, 60, 90, "ELC..02", ""),
        ("dsp1", 750, 60, 110, "DSPTRN", "mode 1"),
        ("e03", 890, 60, 90, "ELC..03", ""),
        ("demand", 1010, 10, 110, "Demand", ""),
        ("export", 1010, 110, 110, "Export", "interconnector"),
        ("import", 10, 210, 140, "Import", "interconnector"),
        ("e04", 180, 210, 90, "ELC..04", ""),
        ("dsp2", 300, 210, 110, "DSPTRN", "mode 2"),
    ]
    arrows = [
        (130, 50, 160, 30), (130, 50, 160, 130), (300, 30, 330, 30), (300, 130, 330, 130),
        (420, 30, 450, 30), (420, 130, 450, 130), (600, 30, 630, 80), (600, 130, 630, 80),
        (720, 80, 750, 80), (860, 80, 890, 80), (980, 80, 1010, 30), (980, 80, 1010, 130),
        (150, 230, 180, 230), (270, 230, 300, 230), (410, 230, 375, 150),
    ]
    svg = ['<svg class="chain" viewBox="0 0 1130 260" role="img" aria-label="The chain inside one node, from supply to demand">',
           '<defs><marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">'
           '<path d="M0 0L10 5L0 10z" class="ah"/></marker></defs>']
    for x1, y1, x2, y2 in arrows:
        svg.append(f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" class="arrow" marker-end="url(#ah)"/>')
    for key, x, y, w, line1, line2 in boxes:
        kind = "fuel" if key.startswith("e0") else ("grid" if "grid" in key else "tech")
        svg.append(
            f'<g class="box {kind}" tabindex="0" data-note="{key}"><rect x="{x}" y="{y}" width="{w}" height="40" rx="8"/>'
            f'<text x="{x + w / 2}" y="{y + (17 if line2 else 26)}">{e(line1)}</text>'
            + (f'<text class="sub" x="{x + w / 2}" y="{y + 34}">{e(line2)}</text>' if line2 else "")
            + "</g>"
        )
    svg.append("</svg>")
    note_html = "".join(f'<div class="note-item" data-note-for="{k}" hidden>{v}</div>' for k, v in notes.items())
    return (
        f'<figure class="hover-diagram" data-hover>{"".join(svg)}'
        f'<div class="note-panel" aria-live="polite"><div class="note-item" data-note-for="" >Hover over or tap a box to read what it does.</div>{note_html}</div>'
        f'<figcaption>One node of the {e(app_config(site, setup_id)["name"])} model. Schematic, not model output. The loss is read from '
        f'{_source_link(site, site.setups[setup_id].sources["country_config"])}.</figcaption></figure>'
    )


# Transmission depth -----------------------------------------------------------------


def block_families_table(site, page, argument):
    out = []
    for setup_id in applications(site):
        info = site.setups[setup_id]
        rows = []
        for category in INTERNAL_CATEGORIES:
            techs = info.families.get(category, [])
            if not techs:
                continue
            carries, tier = CATEGORY_TEXT[category]
            rows.append([_code(techs[0][:6] + "<node>"), e(category), e(carries), _code(f"ELC<node>{tier}"),
                         _code("ELC<node>02"), str(len(techs))])
        out.append(f'<h3 id="families-{e(setup_id)}">{e(app_config(site, setup_id)["name"])}</h3>')
        out.append(_table(["Family", "Taxonomy category", "Carries", "From", "To", "Technologies"], rows))
        values = [
            ["Internal grid technologies in all", value_span(site, f"{setup_id}.internal_grid_technologies")],
        ]
        for name, label, kind in (
            ("transmission_loss", "Loss inside a node", "pct"),
            ("residual_peak_margin", "Base-year grid as a multiple of the base-year peak", ""),
            ("operational_life", "Life of every internal family, years", ""),
        ):
            if has(site, f"{setup_id}.internal.{name}"):
                values.append([label, value_span(site, f"{setup_id}.internal.{name}", kind)])
        out.append(_table(["Setting", "Value"], values, "kv"))
        out.append(f'<p class="small">Counts from {_source_link(site, info.sources["taxonomy"])}; settings from the '
                   f'<code>internal_transmission</code> block of {_source_link(site, info.sources["country_config"])}. '
                   "The technology count of each row is read from the taxonomy, so the six rows sum to the total.</p>")
    return "".join(out)


def block_corridor_table(site, page, argument):
    out = []
    for setup_id in applications(site):
        info = site.setups[setup_id]
        if not info.corridor_costs:
            out.append(f"<p>{e(app_config(site, setup_id)['name'])}: no corridor cost table in its scenario workbook.</p>")
            continue
        rows = []
        for row in info.corridor_costs:
            tech = row["tech"]
            refs = _reference_links(info.corridor_references.get(tech, ""))
            rows.append([
                _code(tech), e(str(row["corridor"])),
                value_span(site, f"{setup_id}.corridor.{tech}.capital"),
                value_span(site, f"{setup_id}.corridor.{tech}.fixed"),
                f"{value_span(site, f'{setup_id}.corridor.{tech}.low')} to {value_span(site, f'{setup_id}.corridor.{tech}.high')}",
                value_span(site, f"{setup_id}.corridor.{tech}.residual") if has(site, f"{setup_id}.corridor.{tech}.residual") else "",
                refs,
            ])
        out.append(f'<h3 id="corridors-{e(setup_id)}">{e(app_config(site, setup_id)["name"])}</h3>')
        out.append(_table(
            ["Technology", "Corridor", f"Capital cost, {e(info.corridor_units)} per kW",
             f"Fixed cost, {e(info.corridor_units)} per kW a year",
             "Low to high capital cost", "Existing capacity in the first year, GW", "Sources"], rows, "wide"))
        out.append(f'<p class="small">Read from the <code>Transmission_Cost_Basis</code> and <code>Interconnector_Params</code> sheets of '
                   f'{_source_link(site, info.sources["scenario_workbook"])}. The low and high values feed the cost bounding scenarios.</p>')
    return "".join(out)


_URL = re.compile(r"https?://[^\s,;)\]]+[^\s,;.)\]]")


def _reference_links(text: str) -> str:
    urls = []
    for url in _URL.findall(text):
        if url not in urls:
            urls.append(url)
    return " ".join(f'<a href="{e(url)}" rel="noopener">[{n}]</a>' for n, url in enumerate(urls, 1))


def block_import_caps(site, page, argument):
    out = []
    for setup_id in applications(site):
        info = site.setups[setup_id]
        name = app_config(site, setup_id)["name"]
        out.append(f'<h3 id="caps-{e(setup_id)}">{e(name)}</h3>')
        cap = info.link_import_cap
        if cap:
            links = ", ".join(_code(link) for link in cap.get("links") or [])
            roots = ", ".join(_code(r) for r in info.historical_trade_roots)
            out.append(
                f"<p><strong>Per-link policy.</strong> From {value_span(site, f'{setup_id}.link_import_cap.first_year')}, the flow from "
                f"{_code(cap.get('from_country', ''))} to {_code(cap.get('to_country', ''))} on {links} stays within what the existing "
                f"links carry, in the scenarios built from the roots with historical trade constraints ({roots}). "
                f"Set in {_source_link(site, info.manifest)} under <code>policies.link_import_cap</code>.</p>"
            )
        gross = [d for d in info.derived if d["gross_import_cap"]]
        if gross:
            rows = []
            for d in gross:
                g = d["gross_import_cap"]
                rows.append([
                    _code(d["name"]), e(", ".join(g.get("countries") or [])),
                    value_span(site, f"{setup_id}.gross_import_cap.{d['name']}.fraction", "pct"),
                    value_span(site, f"{setup_id}.gross_import_cap.{d['name']}.first_year") + " to "
                    + value_span(site, f"{setup_id}.gross_import_cap.{d['name']}.last_year"),
                    _source_link(site, d["source"]),
                ])
            out.append("<p><strong>Gross import caps.</strong> Scenarios with a <code>gross_import_cap</code> block:</p>")
            out.append(_table(["Scenario", "Countries", "Share of annual demand", "Years", "Scenario file"], rows))
        shares = [d for d in info.derived if "import_share_cap" in _edit_ops(site, d)]
        if shares:
            out.append("<p><strong>Per-corridor caps.</strong> Scenarios whose edits use <code>op: import_share_cap</code>: "
                       + ", ".join(_code(d["name"]) for d in shares) + ".</p>")
        if not cap and not gross and not shares:
            out.append("<p>No import cap is set in this application.</p>")
    return "".join(out)


def _edit_ops(site, derived: dict) -> set[str]:
    from ostram.pipeline.scenarios.registry import load_scenario_file

    block = load_scenario_file(site.paths.project_root / derived["source"])
    return {str(edit.get("op", "")) for edit in block.get("edits") or []}


def block_internal_cost_calc(site, page, argument):
    out = []
    for setup_id in applications(site):
        if not has(site, f"{setup_id}.internal.base_capital_cost"):
            continue
        v = lambda name, kind="": value_span(site, f"{setup_id}.internal.{name}", kind)  # noqa: E731
        out.append(f'<h3 id="calc-{e(setup_id)}">{e(app_config(site, setup_id)["name"])}</h3>')
        steps = [
            ("Base capital cost, USD per kW, cost year",
             f"{v('esmap_capital_cost_usd2010_per_kw')} &times; {v('muv_2023')} / {v('muv_2010')}", v("base_capital_cost", "f3")),
            ("Base fixed cost, USD per kW a year, cost year",
             f"{v('fixed_cost_share')} &times; {v('base_capital_cost', 'f3')}", v("base_fixed_cost", "f4")),
            ("Renewable add-on, USD per kW, cost year",
             f"{v('gec2_cost_rs_crore')} &times; 10<sup>7</sup> / ({v('gec2_capacity_mw')} &times; 10<sup>3</sup>) / {v('inr_per_usd_2022')} &times; {v('muv_2023')} / {v('muv_2022')}",
             v("re_add_on", "f3")),
            ("Renewable multiplier", "(base + add-on) / base", v("re_multiplier", "f4")),
            ("Renewable capital cost, USD per kW, cost year", "base &times; multiplier", v("re_capital_cost", "f3")),
        ]
        rows = [[label, formula, result] for label, formula, result in steps]
        out.append(_table(["Quantity", "Formula with the released inputs", "Result"], rows, "calc"))
        out.append(f'<p class="small">Every input is a key of the <code>internal_transmission</code> block of '
                   f'{_source_link(site, site.setups[setup_id].sources["country_config"])}; the results are computed '
                   "when the site is built, with the formula of the internal transmission rule. The rule rounds the "
                   "compiled values; the guide below shows them.</p>")
    return "".join(out)


def block_repowering(site, page, argument):
    out = []
    for setup_id in applications(site):
        info = site.setups[setup_id]
        rows = []
        for category in ("TRANSMISSION_REPOT_RE", "TRANSMISSION_REPOT_NR", "TRANSMISSION_UNPLANNED_RE", "TRANSMISSION_UNPLANNED_NR"):
            techs = info.families.get(category, [])
            if techs:
                rows.append([_code(techs[0][:6] + "<node>"), e(CATEGORY_TEXT[category][0]), str(len(techs))])
        out.append(f'<h3 id="rpo-{e(setup_id)}">{e(app_config(site, setup_id)["name"])}</h3>')
        out.append(_table(["Family", "Carries", "Technologies"], rows) if rows else "<p>No repowering or new-line families.</p>")
    return "".join(out)


# Scenarios ----------------------------------------------------------------------------


def block_scenarios(site, page, argument):
    out = []
    targets = [page.setup] if page.setup else applications(site)
    for setup_id in targets:
        info = site.setups[setup_id]
        name = app_config(site, setup_id)["name"]
        if not page.setup:
            out.append(f'<h3 id="scenarios-{e(setup_id)}">{e(name)}</h3>')
        out.append(
            f"<p>{value_span(site, f'{setup_id}.scenarios')} scenarios: {value_span(site, f'{setup_id}.support_scenarios')} support "
            f"and {value_span(site, f'{setup_id}.decision_scenarios')} decision scenarios. Of the decision scenarios, "
            f"{value_span(site, f'{setup_id}.derived_scenarios')} are derived from a root by scenario edits. Registry: "
            f"{_source_link(site, info.sources['scenario_registry'])}.</p>"
        )
        lines = ["flowchart LR"]
        for root in info.roots:
            lines.append(f'    {root["name"]}["{root["name"]}<br/>{root["role"]} root"]')
        for root in info.roots:
            for dep in root["dependencies"]:
                kind, _, parent = dep.partition(": ")
                lines.append(f'    {parent} -->|"{kind}"| {root["name"]}')
        groups: dict[str, int] = {}
        for derived in info.derived:
            groups[derived["base"]] = groups.get(derived["base"], 0) + 1
        for base, count in groups.items():
            lines.append(f'    {base} -.->|"edits"| D_{base}["{count} derived scenarios"]')
        out.append(f'<pre class="mermaid">{e(chr(10).join(lines))}</pre>')
        rows = [[_code(r["name"]), e(r["role"]), e("; ".join(r["dependencies"]) or "none")] for r in info.roots]
        out.append(_table(["Root", "Role", "Depends on"], rows))
        rows = [[_code(d["name"]), _code(d["base"]), str(d["edits"]) + (f" after those of {_code(d['reuse'])}" if d["reuse"] else ""),
                 e(d["description"])] for d in info.derived]
        if rows:
            out.append(_table(["Derived scenario", "Built from", "Edits", "What it changes"], rows, "wide"))
            out.append(f'<p class="small">Descriptions are the <code>description</code> field of each '
                       f'<code>scenario.yaml</code> under {_code(posixpath.dirname(posixpath.dirname(info.derived[0]["source"])) + "/")}.</p>')
    return "".join(out)


# Running the model ----------------------------------------------------------------------


def block_cli(site, page, argument):
    main = importlib.import_module("ostram.__main__")
    rows = []
    for name, route in list(main.ROUTES.items()) + list(main.EXTENDED_ROUTES.items()):
        rows.append([_code(route.program), e(route.help), _code(route.module_name)])
    parser = main.build_parser()
    for action in parser._subparsers._group_actions:  # noqa: SLF001
        for choice in action._choices_actions:  # noqa: SLF001
            if choice.dest not in main.ROUTES and choice.dest not in main.EXTENDED_ROUTES:
                rows.append([_code(f"python -m ostram {choice.dest}"), e(choice.help or ""), _code("ostram.__main__")])
    options = []
    for action in parser._actions:  # noqa: SLF001
        if action.option_strings and action.dest != "help":
            options.append([_code(", ".join(action.option_strings)), e(action.help or "")])
    return (
        _table(["Command", "What it does", "Module"], rows)
        + "<p>Options placed before the command:</p>"
        + _table(["Option", "Meaning"], options)
        + f'<p class="small">Generated from <code>ROUTES</code>, <code>EXTENDED_ROUTES</code> and the parser in '
        f'{_source_link(site, "ostram/__main__.py")}.</p>'
    )


# Pipeline and code map ----------------------------------------------------------------------


def block_code_map(site, page, argument):
    v = lambda name: value_span(site, f"repo.{name}", "int")  # noqa: E731
    rows = []
    for stage in site.config["stages"]:
        rows.append([e(stage["stage"]), e(stage["does"]), _source_link(site, stage["folder"]),
                     v(f"stage.{stage['stage']}.files"), v(f"stage.{stage['stage']}.lines")])
    out = [_table(["Stage", "What it does", "Code", "Python files", "Lines"], rows)]
    rows = [[_source_link(site, row["folder"]), v(f"code.{row['folder']}.files"), v(f"code.{row['folder']}.lines")]
            for row in code_counts(site.paths)]
    out.append('<h3 id="every-package-folder">Every package folder</h3>')
    out.append(_table(["Folder", "Python files", "Lines"], rows))
    summary = [
        ["Test files (<code>tests/**/test_*.py</code>)", v("tests")],
        ["Docs pages (<code>docs/**/*.md</code>)", v("docs_pages")],
        ["Rule modules (<code>ostram/pipeline/scenarios/rules/</code>)", v("rule_modules")],
        ["Registered setups (<code>config/setups/registry.json</code>)", v("setups")],
        ["GitHub Actions workflows (<code>.github/workflows/</code>)", v("workflows")],
    ]
    out.append('<h3 id="repository-counts">Repository counts</h3>')
    out.append(_table(["What", "Count"], summary, "kv"))
    out.append('<p class="small">Counted when the site is built, by walking the repository.</p>')
    return "".join(out)


# Data and sources --------------------------------------------------------------------------


def block_sources(site, page, argument):
    out = []
    for setup_id in applications(site):
        info = site.setups[setup_id]
        name = app_config(site, setup_id)["name"]
        out.append(f'<h3 id="sources-{e(setup_id)}">{e(name)}</h3>')
        out.append(f"<p>The inputs of this application, as its manifest {_source_link(site, info.manifest)} names them:</p>")
        rows = [[_code(role), _code(path)] for role, path in _authority_rows(site, setup_id)]
        out.append(_table(["Role", "File or folder"], rows))
        if info.corridor_references:
            rows = []
            for tech in sorted(info.corridor_references):
                text = info.corridor_references[tech]
                rows.append([_code(tech), _linkify(text)])
            out.append("<p>Sources of the interconnector costs, from the <code>IEEE Reference</code> column of the "
                       "<code>Interconnector_Params</code> sheet:</p>")
            out.append(_table(["Technology", "References"], rows, "refs"))
    return "".join(out)


def _linkify(text: str) -> str:
    parts = []
    last = 0
    for match in _URL.finditer(text):
        parts.append(e(text[last:match.start()]))
        url = match.group(0)
        parts.append(f'<a href="{e(url)}" rel="noopener">{e(url)}</a>')
        last = match.end()
    parts.append(e(text[last:]))
    return "".join(parts)


def block_values_table(site, page, argument):
    rows = []
    for setup_id, info in site.setups.items():
        for key, item in sorted(info.values.items()):
            rows.append([_code(key), value_span(site, key), entry_html(site, key), _code(item.source), e(item.note)])
    for key, item in sorted(site.repo.items()):
        rows.append([_code(key), value_span(site, key), entry_html(site, key), _code(item.source), e(item.note)])
    return _table(["Key", "Value", "Source", "Read from", "Field"], rows, "values")


def render(name: str, argument: str, *, site, page) -> str:
    """Return the HTML of block ``name``."""

    function = globals().get(f"block_{name.replace('-', '_')}")
    if function is None:
        raise KeyError(f"unknown site block {name!r}")
    return function(site, page, argument)


__all__ = ["render", "format_value", "value_span", "entry_html", "page_sources"]

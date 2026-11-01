"""Build the OSTRAM website: ``python -m ostram site --out <folder>``.

The builder reads the page templates in ``ostram/site/pages``, the prose in
``docs/``, every setup in ``config/setups/registry.json`` and their authority
files, and writes plain HTML, CSS and JavaScript. Output is byte-identical for
the same repository; nothing in it depends on the clock or the machine.

Template directives, replaced at build time:

- ``{{v:<key>}}`` or ``{{v:<key>|<format>}}``: one value from :mod:`ostram.site.values`.
  ``@`` in a key stands for the setup of a per-setup page.
- ``{{doc:<path>}}`` or ``{{doc:<path>#<anchor>}}``: a docs page, or one section of it.
- ``{{block:<name>}}`` or ``{{block:<name> <argument>}}``: generated HTML from :mod:`ostram.site.blocks`.

A link written ``href="repo:<path>"`` points to that repository file at the commit
being built. Every value carries the source entry that :mod:`ostram.site.sources`
reads from the repository's records, and each page that prints values ends with
the sources it used.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
import html
import json
from pathlib import Path
import posixpath
import re
import shutil
import sys
from typing import Sequence

from ostram import yaml_io
from ostram.paths import ProjectPaths, resolve_paths
from ostram.site import blocks, sources
from ostram.site.blocks import value_span
from ostram.site.markdown import MarkdownRenderer, headings_of, section, slugify
from ostram.site.values import SetupInfo, registered_setups, repository_values, setup_info

PACKAGE = Path(__file__).resolve().parent
MARKER = ".ostram-site"
_DIRECTIVE = re.compile(r"\{\{(v|doc|block):([^}]*)\}\}")
_HEADER = re.compile(r"\A---\n(.*?)\n---\n", re.DOTALL)
_VOID = {"br", "hr", "img", "input", "meta", "link", "source", "wbr", "col", "area", "base", "embed", "track"}
_EXCLUDED = {"a", "code", "pre", "h1", "h2", "h3", "h4", "h5", "h6", "svg", "script", "style", "button", "th", "summary", "label"}


@dataclass
class Page:
    """One page of the site."""

    path: str
    template: str
    title: str
    section: str
    summary: str = ""
    setup: str | None = None
    body: str = ""
    headings: list[tuple[int, str, str]] = field(default_factory=list)


@dataclass
class Site:
    """Everything a page needs while it renders."""

    paths: ProjectPaths
    config: dict
    setups: dict[str, SetupInfo]
    pages: list[Page]
    doc_titles: dict[str, str]
    anchors: dict[tuple[str, str], str] = field(default_factory=dict)
    whole_docs: dict[str, str] = field(default_factory=dict)
    glossary: list[tuple[str, str]] = field(default_factory=list)
    uses_mermaid: set[str] = field(default_factory=set)
    repo: dict = field(default_factory=dict)
    commit: str | None = None
    entries: dict = field(default_factory=dict)
    left_out: dict = field(default_factory=dict)
    page_units: list = field(default_factory=list)

    def blob(self, path: str) -> str:
        """Return the address of a repository file at the commit being built, or on ``main`` without Git."""

        return f"{self.config['repository_url']}/blob/{self.commit or 'main'}/{path}"

    @property
    def docs_root(self) -> Path:
        return self.paths.project_root / "docs"

    def value(self, key: str):
        if key in self.repo:
            return self.repo[key]
        setup_id, _, name = key.partition(".")
        info = self.setups.get(setup_id)
        if info is None or key not in info.values:
            raise KeyError(f"no site value {key!r}")
        return info.values[key]


def rel(from_page: str, to_page: str) -> str:
    """Return the relative address of ``to_page`` seen from ``from_page``."""

    start = posixpath.dirname(from_page) or "."
    return posixpath.relpath(to_page, start)


def _read_template(name: str) -> tuple[dict, str]:
    text = (PACKAGE / "pages" / name).read_text(encoding="utf-8").replace("\r\n", "\n")
    match = _HEADER.match(text)
    header = yaml_io.safe_load(match.group(1)) if match else {}
    return header or {}, text[match.end():] if match else text


def _page_list(site_config: dict, setups: dict[str, SetupInfo]) -> list[Page]:
    pages = []
    for entry in site_config["pages"]:
        if entry.get("per_setup"):
            for setup_id in _application_order(site_config, setups):
                name = _setup_name(site_config, setup_id, setups)
                template = entry["template"] if setup_id in setups else entry["placeholder"]
                pages.append(
                    Page(
                        path=entry["path"].format(setup=setup_id),
                        template=template,
                        title=name,
                        section=entry["section"],
                        setup=setup_id,
                    )
                )
            continue
        header, _ = _read_template(entry["template"])
        pages.append(
            Page(
                path=entry["path"],
                template=entry["template"],
                title=header.get("title", entry["path"]),
                section=entry["section"],
                summary=header.get("summary", ""),
            )
        )
    return pages


def _application_order(site_config: dict, setups: dict[str, SetupInfo]) -> list[str]:
    named = [app["setup"] for app in site_config["applications"]]
    return named + [setup_id for setup_id in setups if setup_id not in named]


def _setup_name(site_config: dict, setup_id: str, setups: dict[str, SetupInfo]) -> str:
    for app in site_config["applications"]:
        if app["setup"] == setup_id:
            return app["name"]
    return setups[setup_id].title


def _doc_titles(docs_root: Path) -> dict[str, str]:
    titles = {}
    for path in sorted(docs_root.rglob("*.md")):
        relative = path.relative_to(docs_root).as_posix()
        headings = headings_of(path.read_text(encoding="utf-8"))
        titles[relative] = headings[0].text if headings else relative
    return titles


def _embeds(text: str) -> list[tuple[str, str | None]]:
    found = []
    for kind, argument in _DIRECTIVE.findall(text):
        if kind == "doc":
            doc, _, anchor = argument.strip().rstrip("^").partition("#")
            found.append((doc, anchor or None))
    return found


def _register_anchors(site: Site) -> None:
    """Record which page holds each embedded docs heading."""

    for page in site.pages:
        _, text = _read_template(page.template)
        for doc, anchor in _embeds(text):
            source = (site.docs_root / doc).read_text(encoding="utf-8")
            part = section(source, anchor) if anchor else source
            for heading in headings_of(part):
                if heading.slug:
                    key = (doc, heading.slug)
                    if key in site.anchors and site.anchors[key] != page.path:
                        raise ValueError(f"{doc}#{heading.slug} is embedded twice")
                    site.anchors[key] = page.path
            site.whole_docs.setdefault(doc, page.path)


def _resolver(site: Site, page: Page, doc: str):
    def resolve(href: str) -> str:
        target, _, anchor = href.partition("#")
        target_doc = posixpath.normpath(posixpath.join(posixpath.dirname(doc), target)) if target else doc
        if anchor and (target_doc, anchor) in site.anchors:
            return _local(site.anchors[(target_doc, anchor)], anchor)
        if not anchor and target_doc in site.whole_docs:
            return rel(page.path, site.whole_docs[target_doc])
        stem = target_doc[:-3] if target_doc.endswith(".md") else target_doc
        return site.config["docs_url"] + stem + ".html" + (f"#{anchor}" if anchor else "")

    def _local(target_page: str, anchor: str) -> str:
        prefix = "" if target_page == page.path else rel(page.path, target_page)
        return f"{prefix}#{anchor}"

    return resolve


def _render_doc(site: Site, page: Page, argument: str) -> str:
    argument = argument.strip()
    drop_title = argument.endswith("^")
    doc, _, anchor = argument.rstrip("^").partition("#")
    source = (site.docs_root / doc).read_text(encoding="utf-8")
    text = section(source, anchor) if anchor else source
    if drop_title:
        text = re.sub(r"\A# .*\n", "", text)
    level = 1 if drop_title else (headings_of(text)[0].level if headings_of(text) else 1)
    renderer = MarkdownRenderer(
        resolve=_resolver(site, page, doc),
        doc_title=lambda target: site.doc_titles.get(f"{target}.md", target),
        shift=2 - level,
    )
    rendered = renderer.render(text)
    if 'class="mermaid"' in rendered.html:
        site.uses_mermaid.add(page.path)
    origin = site.blob("docs/" + doc)
    # A docs link to a repository file on main points to the commit being built.
    body = re.sub(
        r'href="' + re.escape(site.config["repository_url"]) + r'/(?:blob|tree)/main/([^"]*)"',
        lambda m: f'href="{html.escape(site.blob(m.group(1)))}"',
        rendered.html,
    )
    return (
        f'<div class="doc" data-doc="{html.escape(doc)}">'
        f'<p class="doc-origin">From <a href="{origin}" rel="noopener">docs/{html.escape(doc)}</a>'
        f"{' (one section)' if anchor else ''}.</p>{body}</div>"
    )


def _render_body(site: Site, page: Page) -> str:
    _, text = _read_template(page.template)

    def replace(match: re.Match) -> str:
        kind, argument = match.group(1), match.group(2).strip()
        if kind == "v":
            key, _, fmt = argument.partition("|")
            if key.startswith("@."):
                key = f"{page.setup}{key[1:]}"
            return value_span(site, key, fmt)
        if kind == "doc":
            return _render_doc(site, page, argument)
        name, _, rest = argument.partition(" ")
        return blocks.render(name, rest.strip(), site=site, page=page)

    body = _DIRECTIVE.sub(replace, text)
    body = _SITE_LINK.sub(lambda m: f'{m.group(1)}="{_site_href(page.path, m.group(2))}"', body)
    body = _REPO_LINK.sub(lambda m: f'href="{html.escape(site.blob(m.group(1)))}"', body)
    if 'class="mermaid"' in body:
        site.uses_mermaid.add(page.path)
    return body


_SITE_LINK = re.compile(r'(href|src)="~/([^"]*)"')
_REPO_LINK = re.compile(r'href="repo:([^"]*)"')


def _site_href(from_page: str, target: str) -> str:
    """Return a template link written from the site root (``~/path#anchor``) relative to ``from_page``."""

    path, _, anchor = target.partition("#")
    if path == from_page:
        return f"#{anchor}" if anchor else rel(from_page, path)
    return rel(from_page, path) + (f"#{anchor}" if anchor else "")


def _glossary_terms(site: Site) -> list[tuple[str, str]]:
    text = (site.docs_root / "glossary.md").read_text(encoding="utf-8")
    terms = []
    for line in text.split("\n## ", 1)[0].splitlines():
        match = re.match(r"^\|\s*([^|]+?)\s*\|", line)
        if match and match.group(1) not in ("Term", "---") and not set(match.group(1)) <= set("-: "):
            term = match.group(1).replace("`", "")
            terms.append((term, "g-" + slugify(term)))
    return terms


def _link_terms(site: Site, page: Page, body: str) -> str:
    """Link the first use of each glossary term on a page to the glossary."""

    if page.path == site.config["glossary_page"]:
        return body
    glossary = rel(page.path, site.config["glossary_page"])
    remaining = {term.lower(): anchor for term, anchor in site.glossary}
    pattern = re.compile(
        r"(?<![\w-])("
        + "|".join(re.escape(term) for term in sorted(remaining, key=len, reverse=True))
        + r")(?![\w-])",
        re.IGNORECASE,
    )
    parts = re.split(r"(<[^>]+>)", body)
    stack: list[bool] = []
    for index, part in enumerate(parts):
        if part.startswith("<"):
            match = re.match(r"<(/?)([a-zA-Z0-9]+)", part)
            if not match or part.endswith("/>") or match.group(2).lower() in _VOID:
                continue
            if match.group(1):
                if stack:
                    stack.pop()
            else:
                stack.append(match.group(2).lower() in _EXCLUDED or 'class="eyebrow"' in part)
            continue
        if any(stack) or not part.strip() or not remaining:
            continue

        def link(match: re.Match) -> str:
            anchor = remaining.pop(match.group(1).lower(), None)
            if anchor is None:
                return match.group(1)
            return f'<a class="term" href="{glossary}#{anchor}">{match.group(1)}</a>'

        parts[index] = pattern.sub(link, part)
    return "".join(parts)


def _glossary_ids(site: Site, body: str) -> str:
    """Give each glossary row an anchor."""

    anchors = {term: anchor for term, anchor in site.glossary}

    def add(match: re.Match) -> str:
        cell = re.sub(r"<[^>]+>", "", match.group(1))
        cell = html.unescape(cell)
        anchor = anchors.get(cell)
        return f'<tr id="{anchor}">\n<td>{match.group(1)}</td>' if anchor else match.group(0)

    head, separator, rest = body.partition("</table>")
    return re.sub(r"<tr>\n<td>(.*?)</td>", add, head) + separator + rest


def _headings(body: str) -> list[tuple[int, str, str]]:
    found = []
    for level, anchor, text in re.findall(r'<h([23]) id="([^"]+)"[^>]*>(.*?)</h\1>', body, re.DOTALL):
        found.append((int(level), anchor, html.unescape(re.sub(r"<[^>]+>", "", text)).strip()))
    return found


def _nav(site: Site, page: Page) -> str:
    groups: dict[str, list[Page]] = {}
    for item in site.pages:
        groups.setdefault(item.section, []).append(item)
    out = ['<nav class="sidenav" aria-label="Site">']
    for name in site.config["sections"]:
        items = groups.get(name, [])
        if not items:
            continue
        out.append(f'<p class="sidenav-group">{html.escape(name)}</p><ul>')
        for item in items:
            current = ' aria-current="page"' if item.path == page.path else ""
            out.append(
                f'<li><a href="{rel(page.path, item.path)}"{current}>{html.escape(item.title)}</a>'
            )
            if item.path == page.path:
                toc = [h for h in page.headings if h[0] == 2]
                if toc:
                    out.append('<ul class="toc">')
                    out += [
                        f'<li><a href="#{anchor}">{html.escape(text)}</a></li>'
                        for _, anchor, text in toc
                    ]
                    out.append("</ul>")
            out.append("</li>")
        out.append("</ul>")
    out.append("</nav>")
    return "".join(out)


def _top_links(site: Site, page: Page) -> str:
    out = []
    for item in site.config["top_nav"]:
        target = item["path"]
        current = ' aria-current="page"' if page.section == item["section"] else ""
        out.append(f'<a class="pill" href="{rel(page.path, target)}"{current}>{html.escape(item["label"])}</a>')
    return "".join(out)


def _layout(site: Site, page: Page, body: str) -> str:
    asset = lambda name: rel(page.path, f"assets/{name}")  # noqa: E731
    title = "OSTRAM" if page.path == "index.html" else f"{page.title} · OSTRAM"
    mermaid = ""
    if page.path in site.uses_mermaid:
        mermaid = (
            '<script type="module">'
            f'import mermaid from "{site.config["mermaid_url"]}";'
            "const dark=document.documentElement.dataset.theme==='dark';"
            "mermaid.initialize({startOnLoad:true,securityLevel:'strict',theme:dark?'dark':'neutral'});"
            "</script>"
        )
    description = html.escape(page.summary or site.config["description"])
    pinned = (
        "" if site.commit else
        " Git was not available when this site was built, so links to repository files point to the main branch."
    )
    return (
        "<!doctype html>\n"
        '<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"<title>{html.escape(title)}</title>\n"
        f'<meta name="description" content="{description}">\n'
        f'<link rel="stylesheet" href="{asset("site.css")}">\n'
        f'<script src="{asset("theme.js")}"></script>\n'
        "</head>\n<body>\n"
        '<a class="skip" href="#main">Skip to content</a>\n'
        '<header class="topbar"><div class="topbar-inner">'
        f'<a class="brand" href="{rel(page.path, "index.html")}"><span class="brand-mark" aria-hidden="true"></span>OSTRAM</a>'
        f'<nav class="pills" aria-label="Sections">{_top_links(site, page)}</nav>'
        '<div class="tools">'
        '<button class="tool search-open" type="button" aria-label="Search (Ctrl K)">'
        '<span aria-hidden="true">&#9906;</span><kbd>Ctrl K</kbd></button>'
        '<button class="tool theme-toggle" type="button" aria-label="Switch light or dark mode">'
        '<span aria-hidden="true">&#9680;</span></button>'
        '<button class="tool menu-toggle" type="button" aria-label="Open the page list" aria-expanded="false">'
        '<span aria-hidden="true">&#9776;</span></button>'
        "</div></div></header>\n"
        '<div class="frame">\n'
        f'<aside class="sidebar">{_nav(site, page)}</aside>\n'
        f'<main id="main" class="content">\n{body}\n</main>\n'
        "</div>\n"
        '<footer class="footer"><p>Generated from the '
        f'<a href="{site.config["repository_url"]}" rel="noopener">OSTRAM repository</a> by '
        "<code>python -m ostram site</code>. Every count and value on this site is read from the "
        "repository when the site is built. Hover over a value for its source; "
        f'<a href="{rel(page.path, "data.html")}#values-on-this-site">'
        f"the values table</a> lists the source and the file of each one.{pinned}</p></footer>\n"
        '<div class="search" hidden><div class="search-box" role="dialog" aria-modal="true" aria-label="Search">'
        '<input type="search" placeholder="Search the site" aria-label="Search the site">'
        '<ol class="search-results"></ol><p class="search-hint">Enter opens the first result. Esc closes.</p>'
        "</div></div>\n"
        f'<script src="{asset("search-index.js")}"></script>\n'
        f'<script src="{asset("site.js")}"></script>\n'
        f"{mermaid}\n"
        "</body>\n</html>\n"
    )


def _search_entries(site: Site, page: Page, body: str) -> list[dict]:
    entries = []
    pieces = re.split(r'(<h[23] id="[^"]+"[^>]*>.*?</h[23]>)', body, flags=re.DOTALL)
    heading, anchor = page.title, ""
    for piece in pieces:
        match = re.match(r'<h[23] id="([^"]+)"[^>]*>(.*?)</h[23]>', piece, re.DOTALL)
        if match:
            anchor = match.group(1)
            heading = html.unescape(re.sub(r"<[^>]+>", "", match.group(2))).strip()
            continue
        text = re.sub(r"<(script|svg|style)[^>]*>.*?</\1>", " ", piece, flags=re.DOTALL)
        text = html.unescape(re.sub(r"<[^>]+>", " ", text))
        text = " ".join(text.split())
        if text or anchor:
            entries.append(
                {
                    "p": page.title,
                    "h": heading,
                    "u": page.path + (f"#{anchor}" if anchor else ""),
                    "t": text[:600],
                }
            )
    return entries


def _write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(text)


def _prepare_out(out: Path) -> None:
    if out.exists():
        if not out.is_dir():
            raise SystemExit(f"--out is a file: {out}")
        if any(out.iterdir()) and not (out / MARKER).is_file():
            raise SystemExit(
                f"--out {out} holds files that this command did not write; give an empty or new folder"
            )
        for child in out.iterdir():
            if child.is_dir() and not child.is_symlink():
                shutil.rmtree(child)
            else:
                child.unlink()
    out.mkdir(parents=True, exist_ok=True)


def build_site(
    out: Path | str,
    *,
    paths: ProjectPaths | None = None,
    registry_path: Path | str | None = None,
) -> list[str]:
    """Build the site into ``out`` and return the page addresses, in order."""

    paths = resolve_paths() if paths is None else paths
    out = Path(out).resolve()
    site_config = yaml_io.load_file(PACKAGE / "content" / "site.yaml")
    setups = {
        manifest.setup_id: setup_info(manifest, paths)
        for manifest in registered_setups(paths, Path(registry_path) if registry_path else None)
    }
    site = Site(
        paths=paths,
        config=site_config,
        setups=setups,
        pages=_page_list(site_config, setups),
        doc_titles=_doc_titles(paths.project_root / "docs"),
    )
    site.repo = repository_values(
        paths,
        site_config["stages"],
        len(setups),
        "config/setups/registry.json" if registry_path is None else Path(registry_path).name,
    )
    site.commit = sources.git_commit(paths.project_root)
    sources.resolve(site)
    site.glossary = _glossary_terms(site)
    _register_anchors(site)

    for page in site.pages:
        if page.setup is not None:
            page.title = _setup_name(site_config, page.setup, setups)
        site.page_units = []
        body = _render_body(site, page)
        body = _glossary_ids(site, body) if page.path == site_config["glossary_page"] else body
        page.body = _link_terms(site, page, body) + blocks.page_sources(site)
        page.headings = _headings(page.body)
        ids = re.findall(r'\sid="([^"]+)"', page.body)
        repeated = sorted({anchor for anchor in ids if ids.count(anchor) > 1})
        if repeated:
            raise ValueError(f"{page.path} has repeated anchors: {repeated}")

    _prepare_out(out)
    search: list[dict] = []
    for page in site.pages:
        _write(out / page.path, _layout(site, page, page.body))
        search += _search_entries(site, page, page.body)
    static = PACKAGE / "static"
    for name in sorted(p.name for p in static.iterdir() if p.is_file()):
        _write(out / "assets" / name, (static / name).read_text(encoding="utf-8").replace("\r\n", "\n"))
    _write(
        out / "assets" / "search-index.js",
        "window.OSTRAM_SEARCH = " + json.dumps(search, ensure_ascii=False, sort_keys=True) + ";\n",
    )
    explainers = PACKAGE / "explainers"
    if explainers.is_dir():
        for source in sorted(explainers.glob("*.html")):
            _write(out / "explainers" / source.name, source.read_text(encoding="utf-8").replace("\r\n", "\n"))
    values = {
        key: {"value": item.value, "source": item.source, "note": item.note, "entry": site.entries[key]}
        for info in setups.values()
        for key, item in sorted(info.values.items())
    }
    values.update({
        key: {"value": item.value, "source": item.source, "note": item.note, "entry": site.entries[key]}
        for key, item in site.repo.items()
    })
    _write(out / "values.json", json.dumps(values, ensure_ascii=False, indent=1, sort_keys=True) + "\n")
    _write(out / ".nojekyll", "")
    _write(out / MARKER, "Written by python -m ostram site. The next build replaces this folder.\n")
    return [page.path for page in site.pages]


def main(argv: Sequence[str] | None = None) -> int:
    """Command line of ``python -m ostram site``."""

    parser = argparse.ArgumentParser(
        prog="python -m ostram site",
        description="Build the OSTRAM website from docs/ and the registered setups.",
    )
    parser.add_argument("--out", required=True, help="Folder to write the site into (replaced on each build).")
    arguments = parser.parse_args(sys.argv[1:] if argv is None else argv)
    pages = build_site(arguments.out)
    print(f"Wrote {len(pages)} pages to {Path(arguments.out).resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

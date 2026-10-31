"""The MyST Markdown subset that ``docs/`` uses, rendered to HTML.

The docs are built by Sphinx with MyST. The site renders the same files with
``markdown-it-py`` (CommonMark plus tables) and handles the MyST additions the
pages use: heading anchors to level 3, ``mermaid`` fences, ``:::{note}`` and
``:::{warning}`` colon fences and the ``{doc}`` role. Anchors follow MyST's
slug rule, so a link such as ``rules.md#build-caps`` finds the same heading.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import html
import re
from typing import Callable

from markdown_it import MarkdownIt

ANCHOR_LEVELS = 3
_FENCE = re.compile(r"^(\s*)(`{3,}|~{3,})")
_COLON_OPEN = re.compile(r"^(:{3,})\{(note|warning|tip|important)\}\s*$")
_HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_DOC_ROLE = re.compile(r"\{doc\}`([^`<]*?)(?:\s*<([^>`]+)>)?`")
_PLACEHOLDER = "OSTRAMBLOCK{:04d}"


def slugify(title: str) -> str:
    """Return the MyST anchor of a heading text."""

    return re.sub(r"[^\w一-鿿\- ]", "", title.strip().lower().replace(" ", "-"))


def plain_heading(text: str) -> str:
    """Return the text of a Markdown heading without inline markup."""

    text = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", text)
    return text.replace("`", "").replace("**", "").replace("*", "")


@dataclass
class Heading:
    level: int
    slug: str
    text: str


@dataclass
class Rendered:
    html: str
    title: str
    headings: list[Heading] = field(default_factory=list)


def _unique(slug: str, used: set[str]) -> str:
    if slug not in used:
        used.add(slug)
        return slug
    index = 1
    while f"{slug}-{index}" in used:
        index += 1
    used.add(f"{slug}-{index}")
    return f"{slug}-{index}"


def headings_of(text: str) -> list[Heading]:
    """Return every heading of a Markdown text, outside code fences, with its anchor."""

    found: list[Heading] = []
    used: set[str] = set()
    fence = None
    for line in text.splitlines():
        match = _FENCE.match(line)
        if match:
            marker = match.group(2)
            if fence is None:
                fence = marker[0] * len(marker)
            elif marker.startswith(fence):
                fence = None
            continue
        if fence is not None:
            continue
        match = _HEADING.match(line)
        if match:
            level = len(match.group(1))
            text_ = plain_heading(match.group(2))
            slug = _unique(slugify(text_), used) if level <= ANCHOR_LEVELS else ""
            found.append(Heading(level, slug, text_))
    return found


def section(text: str, slug: str) -> str:
    """Return the part of ``text`` from the heading ``slug`` to the next heading of its level or above."""

    lines = text.splitlines()
    used: set[str] = set()
    fence = None
    start = level = None
    for number, line in enumerate(lines):
        match = _FENCE.match(line)
        if match:
            marker = match.group(2)
            if fence is None:
                fence = marker[0] * len(marker)
            elif marker.startswith(fence):
                fence = None
            continue
        if fence is not None:
            continue
        match = _HEADING.match(line)
        if not match:
            continue
        this_level = len(match.group(1))
        this_slug = (
            _unique(slugify(plain_heading(match.group(2))), used)
            if this_level <= ANCHOR_LEVELS
            else ""
        )
        if start is None and this_slug == slug:
            start, level = number, this_level
        elif start is not None and this_level <= level:
            return "\n".join(lines[start:number]) + "\n"
    if start is None:
        raise KeyError(f"no heading with anchor {slug!r}")
    return "\n".join(lines[start:]) + "\n"


def _colon_fences(text: str, blocks: list[str], render: Callable[[str], str]) -> str:
    """Replace ``:::{note}`` blocks with placeholders for pre-rendered HTML."""

    lines = text.splitlines()
    out: list[str] = []
    index = 0
    while index < len(lines):
        match = _COLON_OPEN.match(lines[index])
        if not match:
            out.append(lines[index])
            index += 1
            continue
        marker, kind = match.groups()
        body: list[str] = []
        index += 1
        while index < len(lines) and lines[index].strip() != marker:
            body.append(lines[index])
            index += 1
        index += 1
        blocks.append(
            f'<div class="admonition {kind}"><p class="admonition-title">'
            f"{kind.capitalize()}</p>{render(chr(10).join(body))}</div>"
        )
        out += ["", _PLACEHOLDER.format(len(blocks) - 1), ""]
    return "\n".join(out) + "\n"


class MarkdownRenderer:
    """Render docs Markdown to HTML with anchors and rewritten links.

    ``resolve(href)`` returns the site address of a relative link, ``doc_title(target)``
    the title of a docs page named by the ``{doc}`` role.
    """

    def __init__(
        self,
        *,
        resolve: Callable[[str], str],
        doc_title: Callable[[str], str],
        shift: int = 0,
    ) -> None:
        self.resolve = resolve
        self.doc_title = doc_title
        self.shift = shift
        self.md = MarkdownIt("commonmark", {"html": True}).enable("table")
        self.md.add_render_rule("fence", self._fence)

    @staticmethod
    def _fence(renderer, tokens, index, options, env):
        token = tokens[index]
        info = token.info.strip().strip("{}").split()
        language = info[0] if info else ""
        content = html.escape(token.content)
        if language == "mermaid":
            return f'<pre class="mermaid">{content}</pre>\n'
        css = f' class="language-{html.escape(language)}"' if language else ""
        return f"<pre><code{css}>{content}</code></pre>\n"

    def _roles(self, text: str) -> str:
        def replace(match: re.Match) -> str:
            label, target = match.group(1).strip(), match.group(2)
            if target is None:
                target, label = label, ""
            target = target.strip()
            return f"[{label or self.doc_title(target)}]({target}.md)"

        return _DOC_ROLE.sub(replace, text)

    def render(self, text: str) -> Rendered:
        """Return the HTML, the title and the headings of one Markdown text."""

        blocks: list[str] = []
        inner = MarkdownRenderer(resolve=self.resolve, doc_title=self.doc_title)
        text = self._roles(text)
        text = _colon_fences(text, blocks, lambda body: inner.render(body).html)
        tokens = self.md.parse(text)
        used: set[str] = set()
        headings: list[Heading] = []
        title = ""
        for position, token in enumerate(tokens):
            if token.type == "heading_open":
                level = int(token.tag[1])
                content = tokens[position + 1].content
                text_ = plain_heading(content)
                if level == 1 and not title:
                    title = text_
                new_level = min(level + self.shift, 6)
                token.tag = f"h{new_level}"
                tokens[position + 2].tag = f"h{new_level}"
                if level <= ANCHOR_LEVELS:
                    slug = _unique(slugify(text_), used)
                    token.attrSet("id", slug)
                    headings.append(Heading(level, slug, text_))
            if token.type == "inline" and token.children:
                for child in token.children:
                    if child.type == "link_open":
                        href = child.attrGet("href") or ""
                        if not re.match(r"^[a-z][a-z0-9+.-]*:", href):
                            child.attrSet("href", self.resolve(href))
                        else:
                            child.attrSet("rel", "noopener")
        rendered = self.md.renderer.render(tokens, self.md.options, {})
        for number, block in enumerate(blocks):
            rendered = rendered.replace(f"<p>{_PLACEHOLDER.format(number)}</p>", block)
        rendered = rendered.replace("<table>", '<div class="table-wrap"><table>').replace(
            "</table>", "</table></div>"
        )
        return Rendered(rendered, title, headings)

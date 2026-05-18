# -*- coding: utf-8 -*-
"""
build_user_guide.py
===================

Convert ../USER_GUIDE.md into ./USER_GUIDE.html — a single, self-contained
HTML file (CSS embedded) that opens with a double click in any browser.

Mirrors the pattern of ../../../docs/_convert_md_to_pdf.py but produces HTML
(no fpdf2, no font dependency) since the end-user output channel is a browser.

Usage:
    python build_user_guide.py
    python build_user_guide.py --md <path/to/USER_GUIDE.md> --out <path/to/file.html>
"""

from __future__ import annotations

import argparse
import html
from pathlib import Path

import markdown


DOCS_DIR = Path(__file__).resolve().parent
DEFAULT_MD = DOCS_DIR.parent / "USER_GUIDE.md"
DEFAULT_HTML = DOCS_DIR / "USER_GUIDE.html"

# CSS is intentionally embedded so the HTML works as a single file (double-click,
# email-attach, copy-onto-USB). Two design notes:
#   - The `pre` block uses a monospaced family and a soft background. The user
#     guide contains ASCII-art diagrams that depend on monospaced rendering;
#     `white-space: pre` is essential.
#   - Tables get an alternating row color and a sticky header for readability
#     on the larger reference tables.
EMBEDDED_CSS = """
:root {
  --fg: #1f2933;
  --muted: #5b6772;
  --accent: #00477a;
  --accent2: #005b8e;
  --border: #d8dee5;
  --code-bg: #f3f5f7;
  --row-alt: #f0f5fa;
  --link: #0c63a4;
}

html { box-sizing: border-box; }
*, *:before, *:after { box-sizing: inherit; }

body {
  margin: 0;
  padding: 32px 24px 80px;
  background: #fafbfc;
  color: var(--fg);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
               "Helvetica Neue", Arial, sans-serif;
  font-size: 15px;
  line-height: 1.55;
}

main {
  max-width: 880px;
  margin: 0 auto;
  background: #ffffff;
  padding: 32px 40px 64px;
  border: 1px solid var(--border);
  border-radius: 6px;
}

h1, h2, h3, h4 {
  color: var(--accent);
  line-height: 1.25;
  margin-top: 1.6em;
  margin-bottom: 0.4em;
}
h1 { font-size: 26px; border-bottom: 2px solid var(--border); padding-bottom: 8px; }
h2 { font-size: 21px; color: var(--accent2); }
h3 { font-size: 17px; }
h4 { font-size: 15px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.04em; }

p { margin: 0.6em 0; }

a { color: var(--link); text-decoration: none; }
a:hover { text-decoration: underline; }

ul, ol { padding-left: 1.4em; }
li { margin: 0.25em 0; }

code {
  font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
  font-size: 0.92em;
  background: var(--code-bg);
  padding: 1px 5px;
  border-radius: 3px;
  color: #a23a3a;
}

pre {
  font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
  font-size: 13px;
  background: var(--code-bg);
  border: 1px solid var(--border);
  border-radius: 4px;
  padding: 14px 16px;
  overflow-x: auto;
  white-space: pre;
  line-height: 1.4;
}
pre code {
  background: transparent;
  color: inherit;
  padding: 0;
}

table {
  border-collapse: collapse;
  width: 100%;
  margin: 1em 0;
  font-size: 0.95em;
}
th, td {
  border: 1px solid var(--border);
  padding: 6px 10px;
  text-align: left;
  vertical-align: top;
}
th {
  background: var(--accent);
  color: #ffffff;
  font-weight: 600;
}
tr:nth-child(even) td { background: var(--row-alt); }

hr {
  border: 0;
  border-top: 1px solid var(--border);
  margin: 2em 0;
}

blockquote {
  border-left: 4px solid var(--accent2);
  background: #f4f8fc;
  padding: 10px 16px;
  margin: 1em 0;
  color: var(--muted);
}

strong { color: var(--fg); }

/* TOC anchors */
.toc { background: #f4f8fc; border: 1px solid var(--border); padding: 12px 18px; border-radius: 4px; }
.toc ul { padding-left: 1.2em; margin: 0.2em 0; }

footer {
  max-width: 880px;
  margin: 24px auto 0;
  font-size: 12px;
  color: var(--muted);
  text-align: center;
}
"""


def render_html(md_text: str, title: str) -> str:
    """Convert Markdown to a complete HTML document with embedded CSS."""
    body_html = markdown.markdown(
        md_text,
        extensions=[
            "tables",        # GFM-style tables
            "fenced_code",   # ``` code fences
            "toc",           # auto-anchor headings
            "sane_lists",    # better list parsing
        ],
        output_format="html5",
    )
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<style>{EMBEDDED_CSS}</style>
</head>
<body>
<main>
{body_html}
</main>
<footer>OSTRAM &mdash; A3 Multi-Scenario User Guide</footer>
</body>
</html>
"""


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--md", type=Path, default=DEFAULT_MD,
                   help=f"Markdown source (default: {DEFAULT_MD.name} in parent dir)")
    p.add_argument("--out", type=Path, default=DEFAULT_HTML,
                   help=f"HTML output (default: {DEFAULT_HTML.name} in docs/ dir)")
    p.add_argument("--title", default="OSTRAM Multi-Scenario User Guide",
                   help="Title for the HTML document.")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    md_path: Path = args.md
    out_path: Path = args.out
    if not md_path.is_file():
        raise FileNotFoundError(f"Markdown source not found: {md_path}")
    md_text = md_path.read_text(encoding="utf-8")
    html_doc = render_html(md_text, title=args.title)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html_doc, encoding="utf-8")
    print(f"Wrote {out_path} ({out_path.stat().st_size // 1024} KB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

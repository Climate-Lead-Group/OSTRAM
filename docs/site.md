# The website

The OSTRAM website is generated from this repository. `python -m ostram site`
builds plain HTML, CSS and JavaScript from three sources:

- the prose in `docs/`, rendered from the same Markdown files that this
  documentation uses;
- the values of every setup registered in `config/setups/registry.json`, read
  through `ostram/setups.py` from the files that each setup manifest names
  (the interconnector taxonomy, the country and compilation configuration, the
  scenario registry and the scenario workbook);
- the page templates, configuration and static files in `ostram/site/`.

Every count and value on the site is read from those files when the site is
built. Each value carries the file it came from; the site's "Values on this
site" table lists them all, and `values.json` in the output holds the same list.

## Build and preview

From the repository root, in the activated environment:

```powershell
python -m ostram site --out _site
python -m http.server 8000 --directory _site
```

Then open `http://localhost:8000/` in a browser. Opening `_site/index.html`
directly also works. `--out` must name a new folder, an empty folder or a
folder written by an earlier build; the command replaces its contents.
`_site/` is git-ignored: generated HTML is never committed.

The build needs no network. The pages load one external script, Mermaid, from
`cdn.jsdelivr.net` at the version pinned in `ostram/site/content/site.yaml`, to
draw diagrams in the browser.

## How the site is put together

| Path | What it holds |
|---|---|
| `ostram/site/build.py` | The command: page list, layout, docs embedding, glossary links, search index. |
| `ostram/site/values.py` | Every value the site shows, read from the setups and their authority files. |
| `ostram/site/blocks.py` | Generated parts of pages: the worked example, the rules stepper, the node diagram and the tables. |
| `ostram/site/markdown.py` | The MyST subset that `docs/` uses, rendered with `markdown-it-py`. |
| `ostram/site/content/site.yaml` | Page order, navigation, the applications and the Mermaid version. |
| `ostram/site/pages/` | One template per page. |
| `ostram/site/static/` | The stylesheet and scripts. |

A template writes `{{v:<setup>.<key>}}` for a value, `{{doc:<path>}}` or
`{{doc:<path>#<anchor>}}` for a docs page or one section of it, and
`{{block:<name>}}` for a generated part. Links inside an embedded docs page
point to the site page that holds their target, or to this documentation when
the site does not carry it.

## Applications and new setups

The site has one page per registered setup. The applications named in
`site.yaml`, South Asia and Latin America and the Caribbean, also appear on the
home page and the transmission pages. An application that is not registered
yet shows a placeholder; when its setup is added to the registry, the same
templates fill its pages from its files.

## Deployment

`.github/workflows/site.yml` builds the site on every push to `main` and on
manual dispatch, and deploys it to GitHub Pages. The repository's Pages source
must be set to "GitHub Actions" once, by an owner, under Settings, Pages.

## Checks

`tests/website/test_site.py` builds the site twice and compares the output byte
for byte, reads every value back from the setup files, checks every internal
link and anchor and the hosts that pages load, and builds once more with a test
`lac` setup to check that the placeholder pages fill.

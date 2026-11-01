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
built. Each value carries the file it came from and its source entry, described
below; the site's "Values on this site" table lists them all, and `values.json`
in the output holds the same list.

## How sources are shown

Every value on the site shows the source or assumption that the repository
already records for it. The site writes no source of its own: the source map
`ostram/site/content/sources.yaml` holds only the places where each record lives,
and `ostram/site/sources.py` reads the record from there at build time. Each
value gets one entry of four kinds:

| Kind | What the entry holds | Where it is read from |
|---|---|---|
| External source | The citation text and its URL or DOI. Any `ASSUMPTION:` sentence of the same record is shown with it. | The `Basis` column of the `internal_transmission` table in `docs/rules.md`; the comment lines above a key in a YAML file; the `IEEE Reference` (or `Source`) column of the `Interconnector_Params` sheet of the scenario workbook, by scenario, technology and parameter; the tables and reference list of `docs/reference/scenarios/vre_ceiling_provenance.md`. |
| Assumption | The sentence as the repository records it, shown after the word *assumption*, without the leading `ASSUMPTION:`. | The same records. |
| Computed | The formula in words and the keys of its inputs. Each input has its own entry. | The inputs listed under `computed` in the source map. |
| Model setting | A link to the file the value is read from, at the commit being built, labelled "model definition", "scenario setting" or "counted from the repository". | The file of the value. |

A record that names its source only in short, such as "ESMAP (2026)", takes the
URL of the one full citation in the same table or reference list that holds the
same words. A docs table is used only for the file it documents, and only when
its value equals the value in that file; otherwise the build stops.

A value whose record is missing or empty is not printed: the page keeps its
prose without the number, and the value is left out of `values.json`. A value
that the source map does not name stops the build, so a new value needs a
decision about its record first.

On a page:

- the hover on a value shows its entry in one line, then the file it is read from;
- each page that prints values ends with "Sources on this page": one numbered
  entry per distinct source, with external links, assumption sentences in full and
  links to the files of model settings;
- the values table on the Data page has a "Source" column with the same entry.

Every link to a repository file, from entries, embedded docs pages and the code
map, points to `blob/<commit>/<path>`, where the commit is `git rev-parse HEAD` of
the tree being built. Without Git, the links point to `blob/main` and the page
footer says so. A template links to a repository file with `href="repo:<path>"`.

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
| `ostram/site/sources.py` | The source entry of each value, read from the records the source map names. |
| `ostram/site/content/sources.yaml` | The source map: where the record of each value lives. Locations only. |
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
`lac` setup to check that the placeholder pages fill. It also checks that every
printed value has an entry of one of the four kinds, that every external entry
has a URL or DOI, that every assumption sentence is the repository's own, that
every computed entry names inputs with entries, that no entry holds a local path,
a sheet-and-cell reference or the name of an internal document, that values with
no record are left out, and that repository links are pinned to the built
commit, or fall back to `main` without Git.

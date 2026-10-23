# Project Context — MDD-read-py (`mdmreadpy`)

> Last updated: 2026-10-23
> Status: initial capture. Architecture + design decisions + review findings recorded.
>
> **Convention:** read this file at the start of every task and update it at the end,
> so it stays the permanently updated context of the current work (see
> `.clinerules/00-project-context.md`).

## 1. What this package does

`mdmreadpy` reads proprietary MDD files (IBM SPSS Data Collection / Unicom Intelligence
"MDM Document") and produces a JSON "report" that the rest of the global tool
(`mdmreportpy`, the MDD diff tool, Excel export, etc.) consumes.

MDD files are not a plain format you can parse from a public spec: the only supported
reader is the COM API exposed by **IBM/Unicom Professional**. Therefore this package is
a thin, configurable driver over that COM API.

Hard requirements to actually read an MDD:

- **Windows** (COM).
- **IBM/Unicom Professional** installed (`MDM.Document` COM server).
- **pywin32** (`win32com.client`) on Windows.

## 2. Where it lives and how it is built

- Package root: `src/lib/mdmreadpy`
- Source modules: `src/lib/mdmreadpy/src/`
- Vendored submodule (JSON → HTML): `src/lib/mdmreadpy/src/lib/mdmreportpy/`
- Build chain: `build.bat` → version `GENERATED/_VERSION.py` → `mdmreportpy` build →
  `pinliner` single-file bundle.

## 3. Modules

| Module | Role |
| --- | --- |
| `src/read_mdd.py` | `MDMDocument` (open/join, traversal, JSON assembly) + `entry_point` CLI. |
| `src/win32com_adapter.py` | Lazy `win32com.client` import; `Win32Client.Dispatch`; fail-early errors for non-Windows / missing pywin32. |
| `src/launcher.py` | Dispatcher `--program read_mdd|report|report_html|test|version`. |
| `src/lib/mdmreportpy/` | Submodule: renders this JSON to HTML. |

## 4. Output JSON format

`MDMDocument.read(config)` returns a dict (saved as JSON by `entry_point`):

```text
{
  report_type: "MDD",
  source_file,
  report_datetime_utc,
  report_datetime_local,
  source_file_metadata: [ {name, value, role?} ],
  report_scheme: { columns, column_headers, flags: ["data-type:mdd"] },
  sections: [ {name, content: [row, ...]}, ... ]
}
```

Columns are derived from the `features` config:
`name`, then `label` / `attributes` / `properties` / `scripting`, and `langcode-<lang>`
per language when `translations` is enabled.

Sections (defaults) and what each row carries:

- `languages` — one row per document language: `name`, `label` (the `LongName`).
- `shared_lists` — shared-list entries; references may be expanded (`stepinto`) or kept
  as-is (`stepover`) via `config_sharedlists_listcats_stepinto`.
- `fields` — a leading "zero item" (document-level `properties`/`scripting`) followed
  by flattened field rows. Nested categories/elements/subfields/helper fields become
  additional rows with dotted `name` prefixes (e.g. `Field.categories[Cat]`,
  `Block.SubField`).
- `pages` — page rows (and their entries).
- `routing` — one row carrying the routing script (`name` + `label`).

Cell values are flat: `attributes` and `properties` are property lists
(`[{"name": ..., "value": ...}, ...]`); `scripting` and `label` are strings.

## 5. Design decisions

- **COM access is isolated** in `win32com_adapter.py`, so the rest of the code never
  imports `win32com` directly. The import is lazy: the module can be imported on
  non-Windows machines for inspection/testing, and fails early with a readable message
  only when an MDD is actually opened.
- **"open" and "join" both supported** (`MDM.Document.Open` vs `.Join`):
    - `open` is simpler and usually keeps the normal label representation in the
      "Scripts" view, but it opens the file exclusively (blocking) and may fail if the
      file is already in use.
    - `join` creates a new empty document and merges the opened document into it. It
      brings in everything the same way, but in the ".Script" view labels often
      disappear (they live in a different context) — only relevant for `.Script`
      outputs, not this reader. It does not require exclusive access, so it is a more
      stable fallback.
    - Default is `open`; `join` is equally legitimate.
- **Config is a plain dict** with defaults (`features`, `sections`, `contexts`); the CLI
  maps `--config-*` arguments onto it. Not over-validated (stays python-ish).
- **Translations are read via the `Question` context.** Translation overlays are often
  trimmed/removed once a project reaches the data-processing/Analysis stage, so the
  `Question` context is the primary (safe) source for `langcode-*` values.
- **Flattening.** The MDM object model is deeply nested (fields → elements →
  categories → subfields → helper fields). The reader flattens it into a flat list of
  rows with dotted names so the JSON has a uniform, table-friendly shape and can be
  rendered directly by `mdmreportpy`.
- **Error handling.** Fail early with readable, context-rich messages (prints the
  section/row being processed before re-raising). The COM document is released in
  `__del__` and via the `with` statement (`__enter__`/`__exit__`).

## 6. Review findings (2026-10-23)

1. **No automated tests** — correctness is verified manually against real MDDs (which
   are proprietary and git-ignored). Any future refactor needs tests; the natural seam
   is a mocked COM boundary (`MDMDocument` over a fake `Win32Client`) plus a golden JSON
   fixture.
2. **`read_mdd.py` is a monolith** (~850 lines) mixing COM traversal, JSON shaping and
   the CLI. Accepted for now; a natural split is (a) COM adapter, (b) document
   traversal → plain dict, (c) CLI. Not urgent.
3. **Dead code** — line 46 (`'label','attributes','properties','translations',...`)
   is a bare tuple expression that does nothing. *(Removed 2026-10-23.)*
4. **Config defaults duplicated** — once in `MDMDocument.__init__` (`config_default`)
   and again in `entry_point`. Keep in sync; consider a single source of truth.
5. **Category traversal relies on MDM internals** (`item.Elements`, `cat.Type`,
   `cat.IsReference`, `ReferenceName` sanitization, `cat.Type` magic numbers). This is
   inherently fragile; the `ErrorItemNotIterable` guard and per-item try/except already
   soften failures. Documenting these object-model assumptions is an open item.
6. **`translations` feature** initializes the cell to `'???'` and only overwrites it in
   a try/except — on COM failure the cell becomes the exception text. Intentional
   (never silently drop a language), but worth a review.

### Applied code improvements (2026-10-23)

- Modernized the remaining `'{x}'.format(x=...)` calls to f-strings across
  `read_mdd.py` (no behavioral change intended).
- Replaced the mutable `config={}` default arguments with `config=None` + normalize in
  `MDMDocument.__init__` and `read`.
- Removed the broken `argparse.ArgumentError(...)` call (it was missing the required
  `message` argument); with `choices=["stepinto","stepover"]` the value is now validated
  by argparse, so the redundant `if/elif/else` collapsed to a single boolean assignment.
- Fixed the `unrecognized section type` error to report `section_name` (it was reporting
  the unset `section_content`).
- Removed the stray `'label','attributes',...` module-level tuple.
- Not yet tested against real MDDs (no Windows / MDM server available); review before
  release.
- Added module/class/method docstrings to `read_mdd.py`, `win32com_adapter.py` and
  `launcher.py`.
- Clarified `open` vs `join` semantics and the `contexts` (Question/Analysis) meaning
  across comments, docstrings and docs.

## 7. Roadmap

1. **Documentation (this file).** *(this step)*
2. **Formalize the JSON output spec** and reconcile it with the `mdmreportpy` input
   spec — the output of this package is the input of that one, so they must stay
   aligned.
3. **Add tests** around the pure JSON-shaping logic by mocking the COM boundary, plus a
   golden JSON fixture.
4. **Split `read_mdd.py`**, remove dead code (#3), centralize config defaults (#4).

## 8. Open questions

- Are there MDDs where `open` fails due to file locking and only `join` works? (`join`
  is the intended fallback when `open` cannot get exclusive access.)
- Exact MDM object-model assumptions (element/category `Type` codes, `ReferenceName`
  sanitization) — document them or guard them more defensively.
- Whether `scripting`/`translations` should ever be omitted vs returning the exception
  text on COM failure.


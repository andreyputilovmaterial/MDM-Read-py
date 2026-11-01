# MDD-read-py — architecture & output JSON format

Stable reference. Living working-memory notes (findings, roadmap, open questions) live
in `docs/project-context.md`, which is updated on every agent run and is not linked
from user-facing docs.

## 1. Where it fits

`mdmreadpy` is one of several **readers** in the global project. It produces the common
JSON "report" shape that is consumed by:

- `mdmreportpy` (the HTML reporter),
- `mdmexcelreportpy` (the Excel reporter),
- `MDM-TranslationOverlayHelper` (the translation overlay helper),
- the diff tool (compares two JSON reports and writes a new JSON report),
- and any other consumer of the common JSON shape.

## 2. Modules

| Module | Role |
| --- | --- |
| `src/read_mdd.py` | `MDMDocument` (open/join, traversal, JSON assembly) + `entry_point` CLI. |
| `src/win32com_adapter.py` | Lazy `win32com.client` import; `Win32Client.Dispatch`; fail-early errors for non-Windows / missing pywin32. |
| `src/launcher.py` | Dispatcher `--program read_mdd\|report\|report_html\|test\|version`. |
| `src/lib/mdmreportpy/` | Submodule: HTML reporter that consumes this JSON. |

## 3. Output JSON format

The master definition of the common JSON "report" shape lives in **`mdmreportpy`**
(`src/lib/mdmreportpy/docs/report-json-format.md`), and the rich-text cell/header
format in `src/lib/mdmreportpy/docs/rich-text-format.md`. This section describes what
`mdmreadpy` produces (a producer-specific view); the `mdmreportpy` spec is the source
of truth for the common shape.

`MDMDocument.read(config)` returns a dict (saved as JSON by `entry_point`):

```text
{
  report_type: "MDD",
  source_file,
  report_datetime_utc,
  report_datetime_local,
  source_file_metadata: [ {name, value, role?} ],
  report_scheme: { columns, column_headers, flags: ["data-type:mdd", "mdmreadpy:version:<version>", "mdmreadpy:config:row-name-convention:<convention>"] },
  sections: [ {name, content: [row, ...]}, ... ]
}
```

Columns are derived from the `features` config: `name`, then `label` / `attributes` /
`properties` / `scripting`, and `langcode-<lang>` per language when `translations` is
enabled.

Sections (defaults) and what each row carries:

- `languages` — one row per document language: `name`, `label` (the `LongName`).
- `shared_lists` — shared-list entries (see "Shared lists" below).
- `fields` — a leading "zero item" (document-level `properties`/`scripting`) followed
  by flattened field rows. Nested categories/elements/subfields/helper fields become
  additional rows with dotted `name` prefixes (e.g. `Field.Categories[Cat]`,
  `Block.SubField`).
- `pages` — page rows (and their entries).
- `routing` — one row carrying the routing script (`name` + `label`).

Cell values are flat: `attributes` and `properties` are property lists
(`[{"name": ..., "value": ...}, ...]`); `scripting` and `label` are strings.

`report_scheme.column_headers` maps each column id to its human-readable title.
**Decision (2026-11-02):** a header value may use the same rich-text structure as a
cell value — a plain string, a list of parts, `{"text": ..., "role": ...}`, or
`{"parts": [...]}` — so consumers can show/hide or style parts by `role`. The MDD
reader still emits plain strings today; consumers are expected to accept the richer
form.

## 4. Config options

The CLI `--config-*` options control **which parts are written to the output JSON**:

| Option | Default | Meaning |
| --- | --- | --- |
| `--config-features` | `label,attributes,properties,translations,scripting` | Which kinds of content to process and include (each becomes one or more columns). |
| `--config-sections` | `languages,shared_lists,fields,pages,routing` | Which document parts to include. |
| `--config-contexts` | `Question,Analysis` | Contexts used when reading labels and custom properties. |
| `--config-sharedlists-listcats` | `stepinto` | `stepinto` or `stepover` (see below). |
| `--config-row-name-convention` | `ap` | `ap` or `mdm` (see "Row name conventions" below). |

Contexts (`--config-contexts`, usually `Question` and `Analysis`) apply to **labels as
well as custom properties**. `Question` usually holds the question text with formatting
(HTML tags); `Analysis` usually holds cleaned labels — often significantly shortened —
for data processing/analysis (SPSS users dislike long strings, variable names and
labels). Translation overlays are read through the `Question` context, since they are
often trimmed or removed once a project reaches the data-processing/Analysis stage.

### Row name conventions

Each row's `name` is its unique identifier. How it is spelled is controlled by
`--config-row-name-convention` (and the reader config key `config_row_name_convention`):

- **`ap`** (default) — compose the name from parent + `.` + child. Categories become
  `Parent.Categories[Cat]`. In `stepover` mode a shared-list reference becomes
  `Parent.Categories[@SL_Name]` (the `@` marks it as a reference; the prefix is
  configurable through `config_row_name_reference_prefix`, but is intentionally **not**
  a CLI flag).
- **`mdm`** (placeholder) — use each item's `FullName` as-is, assuming it already
  contains the full dotted path. A missing `FullName` is an error. Category /
  shared-list handling is still under review.

The convention used is recorded in the report's `flags` as
`mdmreadpy:config:row-name-convention:<convention>`, so consumers can branch on it.

> The `.` separator is significant: the diff tool uses it to group rows hierarchically
> when comparing MDD reports, and other consumers may rely on it too. Only change the
> convention if you keep both sides of a diff (and any consumers) aligned.

## 5. Shared lists: "step into" vs "step over"

Items that have categories/elements (e.g. a question) do not necessarily list them all
individually. When several items share the same categories, the MDD stores them once in
a **shared list**, and each item holds only a single reference to that shared list.

When such a reference is encountered, the reader can:

- **`stepinto`** — follow the reference and expand the shared list into its full list
  of categories/elements (one line per category/element), or
- **`stepover`** — print only one line with the reference to the shared list.

Both modes are useful and both are necessary, depending on what the consumer needs, so
it is left configurable; there is no single "correct" default.

## 6. Design decisions

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
- **Flattening.** The MDM object model is deeply nested (fields → elements →
  categories → subfields → helper fields). The reader flattens it into a flat list of
  rows with dotted names so the JSON has a uniform, table-friendly shape.
- **Error handling.** Fail early with readable, context-rich messages (prints the
  section/row being processed before re-raising). The COM document is released in
  `__del__` and via the `with` statement (`__enter__`/`__exit__`).

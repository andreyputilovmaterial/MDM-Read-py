# MDD-read-py — architecture & output JSON format

Stable reference. Living working-memory notes (findings, roadmap, open questions) live
in `docs/project-context.md`, which is updated on every agent run and is not linked
from user-facing docs.

## 1. Where it fits

`mdmreadpy` is one of several **readers** in the global project. It produces the common
JSON "report" shape that is consumed by:

- `mdmreportpy` (the HTML reporter),
- the diff tool (compares two JSON reports and writes a new JSON report),
- Excel export and any other consumer of the common JSON shape.

## 2. Modules

| Module | Role |
| --- | --- |
| `src/read_mdd.py` | `MDMDocument` (open/join, traversal, JSON assembly) + `entry_point` CLI. |
| `src/win32com_adapter.py` | Lazy `win32com.client` import; `Win32Client.Dispatch`; fail-early errors for non-Windows / missing pywin32. |
| `src/launcher.py` | Dispatcher `--program read_mdd|report|report_html|test|version`. |
| `src/lib/mdmreportpy/` | Submodule: HTML reporter that consumes this JSON. |

## 3. Output JSON format

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

Columns are derived from the `features` config: `name`, then `label` / `attributes` /
`properties` / `scripting`, and `langcode-<lang>` per language when `translations` is
enabled.

Sections (defaults) and what each row carries:

- `languages` — one row per document language: `name`, `label` (the `LongName`).
- `shared_lists` — shared-list entries (see "Shared lists" below).
- `fields` — a leading "zero item" (document-level `properties`/`scripting`) followed
  by flattened field rows. Nested categories/elements/subfields/helper fields become
  additional rows with dotted `name` prefixes (e.g. `Field.categories[Cat]`,
  `Block.SubField`).
- `pages` — page rows (and their entries).
- `routing` — one row carrying the routing script (`name` + `label`).

Cell values are flat: `attributes` and `properties` are property lists
(`[{"name": ..., "value": ...}, ...]`); `scripting` and `label` are strings.

## 4. Config options

The CLI `--config-*` options control **which parts are written to the output JSON**:

| Option | Default | Meaning |
| --- | --- | --- |
| `--config-features` | `label,attributes,properties,translations,scripting` | Which kinds of content to process and include (each becomes one or more columns). |
| `--config-sections` | `languages,shared_lists,fields,pages,routing` | Which document parts to include. |
| `--config-contexts` | `Question,Analysis` | Contexts used when reading custom properties. |
| `--config-sharedlists-listcats` | `stepinto` | `stepinto` or `stepover` (see below). |

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
- **"open" and "join" both supported** (`MDM.Document.Open` vs `.Join`) because both
  are used elsewhere in the ecosystem (FileTrimmer/PrepData use `join`); neither is
  clearly better, so the method is a config/CLI option.
- **Flattening.** The MDM object model is deeply nested (fields → elements →
  categories → subfields → helper fields). The reader flattens it into a flat list of
  rows with dotted names so the JSON has a uniform, table-friendly shape.
- **Error handling.** Fail early with readable, context-rich messages (prints the
  section/row being processed before re-raising). The COM document is released in
  `__del__` and via the `with` statement (`__enter__`/`__exit__`).

# MDD-read-py (`mdmreadpy`)

Reads proprietary MDD files (IBM SPSS Data Collection / Unicom Intelligence "MDM
Document" format) and writes their contents as JSON that the rest of the MDM toolchain
understands.

> MDD files are the proprietary project/dictionary files produced by Unicom software
> (formerly IBM SPSS Data Collection). To read them, this tool drives the COM API of
> IBM/Unicom Professional. Both are required to actually run it:
>
> - **Windows**, and
> - **IBM/Unicom Professional** installed (the `MDM.Document` COM server),
> - plus **pywin32** (`win32com.client`) on Windows.

## What it does

1. Opens an MDD via the `MDM.Document` COM object (two supported methods: `open` or
   `join`).
2. Walks the document and writes a JSON "report" that contains a **configurable set of
   parts** — see the options below.
3. Saves the JSON next to the MDD (same base name, `.json` suffix).

This package is one of several "readers" in the global project. Its JSON output is
consumed by, among others:

- the HTML report generator (`mdmreportpy`, `report_create.py`),
- the diff tool, which compares two JSON reports (e.g. two MDDs, or an MDD against a
  translation) and writes a new JSON report that `mdmreportpy` then renders.

## Requirements

- Windows.
- IBM/Unicom Professional installed (provides `MDM.Document`).
- Python 3 + `pywin32` (`win32com.client`) on Windows.
  - The import is lazy (`src/win32com_adapter.py`), so the module can be imported on
    non-Windows machines, but any attempt to actually open an MDD fails early with a
    readable message.

## Usage

Direct:

```
python src/read_mdd.py --mdd "path\to\file.mdd"
```

Universal launcher:

```
python src/launcher.py --program read_mdd --mdd "path\to\file.mdd"
python src/launcher.py --program report --inpfile "path\to\file.mdd.json"
```

Or edit and run `run.bat` (example workflow: read MDD → JSON, generate HTML, then
delete the temporary JSON).

### Options (`read_mdd.py`)

These options control **which parts are written to the output JSON** — they configure
the output, not how the file is read.

| Option | Default | Meaning |
| --- | --- | --- |
| `--mdd PATH` | (required) | Input MDD file. |
| `--method open\|join` | `open` | How to open the file — see "open vs join" below. |
| `--config-features LIST` | `label,attributes,properties,translations,scripting` | Which kinds of content to process and include in the output (each becomes one or more columns): `label`, `attributes`, `properties`, `translations`, `scripting`. |
| `--config-sections LIST` | `languages,shared_lists,fields,pages,routing` | Which parts of the document to include in the output: `languages`, `shared_lists`, `fields`, `pages`, `routing`. |
| `--config-contexts LIST` | `Question,Analysis` | Contexts used when reading labels and custom properties (see below). |
| `--config-sharedlists-listcats stepinto\|stepover` | `stepinto` | How to treat a field that references a shared list (see below). |

### Shared lists: "step into" vs "step over"

When the reader reaches an item that has categories/elements (e.g. a question), those
categories are usually not all listed on the item itself. If several questions share
the same categories, the MDD stores them once in a **shared list**, and each question
holds just a single reference to that shared list.

When such a reference is encountered, the reader can either:

- **`stepinto`** — follow the reference and expand the shared list into its full list
  of categories/elements (one line per category/element), or
- **`stepover`** — print only one line with the reference to the shared list.

Both modes are useful and both are necessary, depending on what the consumer needs, so
it is left configurable. There is no single "correct" default.

### `open` vs `join`

- **`open`** — the default and the more straightforward mode. It usually keeps the
  normal representation of labels in the "Scripts" view. However, it opens the file
  exclusively (a blocking operation), so it may fail if the MDD is already open.
- **`join`** — creates a new empty document and merges the opened document into it. It
  technically brings in everything (definitions, questions, labels, …), and for a deep
  recursive read with explicit contexts it makes little difference vs `open`. It does
  not require exclusive access, so it is a more stable fallback when `open` fails. One
  caveat: in the ".Script" view labels often disappear with `join` (they live in a
  different context), but that only affects `.Script`-style outputs, not this reader.

### Contexts (`--config-contexts`)

Contexts (usually `Question` and `Analysis`) apply to **labels as well as custom
properties**:

- `Question` usually holds the question text with formatting (HTML tags).
- `Analysis` usually holds cleaned labels for data processing/analysis — often
  significantly shortened, since SPSS users dislike long strings, variable names and
  labels.

Translation overlays are read through the `Question` context: they are often trimmed or
removed once a project reaches the data-processing/Analysis stage, so `Question` is the
primary (safe) source.

## Output JSON shape

```
{
  "report_type": "MDD",
  "source_file": "...",
  "report_datetime_utc": "...",
  "report_datetime_local": "...",
  "source_file_metadata": [ {"name": ..., "value": ..., "role": ...} ],
  "report_scheme": {
    "columns": ["name", "label", "attributes", "properties", "scripting", "langcode-..."],
    "column_headers": {...},
    "flags": ["data-type:mdd"]
  },
  "sections": [
    { "name": "languages",    "content": [...] },
    { "name": "shared_lists", "content": [...] },
    { "name": "fields",       "content": [...] },
    { "name": "pages",        "content": [...] },
    { "name": "routing",      "content": [...] }
  ]
}
```

Each row in `content` is a flat dict keyed by the column ids from
`report_scheme.columns`. `attributes` and `properties` are property lists
(`[{"name": ..., "value": ...}, ...]`). Field rows can be nested — categories,
elements, subfields and helper fields are flattened into additional rows with dotted
`name` prefixes such as `Field.categories[Cat]`. See `docs/architecture.md` for the
full shape and column/row conventions.

## Build

`build.bat` generates `src/GENERATED/_VERSION.py`, rebuilds the vendored
`mdmreportpy`, and produces the single-file `dist/mdmtoolsap_bundle.py`.

## Documentation

- [Architecture & output JSON format](docs/architecture.md) — stable reference for
  developers.


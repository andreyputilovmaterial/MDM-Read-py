# MDD — the input format (principles)

MDD ("metadata document model") is a proprietary Unicom/IBM format used in market
research. This page describes its *principles* — not a full specification. The exact
contracts live in the code, and `mdmreadpy` is the only MDD-specific code in this
repository: `src/read_mdd.py` turns an MDD document into the common JSON report shape.
That JSON (and the rich-text content format inside it) is a separate, toolchain-owned
format with nothing to do with MDD itself — see `mdmreportpy`'s
`docs/report-json-format.md` and `docs/rich-text-format.md`.

## What an MDD file is

On disk an MDD file is XML, but in the host software (IBM/Unicom Professional) it is a
cleanly organized object model. The XML detail is irrelevant to this reader.

## Objects

Everything is organized into **Questions/Fields/Items** (the name varies by context)
and **Categories/Elements**. Each object has:

- a **technical name** (Question Name / Category Name), following the usual
  variable-name rules, and
- a **Label**, which holds the display text.

Labels exist in parallel versions within a single document:

- by **context** — the common ones are `Question` (the live survey shown to
  respondents) and `Analysis` (tables and SPSS);
- by **language overlay**;
- by **versioning**.

## Properties and Modifiers

Each object (Question/Field/Item, Shared list, Page, Category/Element) carries:

- **Properties** — key-value pairs that can also be context-specific (`Question`, `Analysis`, …).
- **Modifiers** — flags that change behavior, for example: randomize the category list,
  pin the element at its position (e.g. last), exclude it from filtering, mark it
  "exclusive" (cannot be combined with other answers), allow routing to continue
  without a response, or not store its value in the case-data database. There are more;
  this list is illustrative.

## Question types

**Scalar:**

- `text`, `date/time`, `long` (integer), `float` (real number);
- **categorical** — single-punch or multi-punch (the most common). The format barely
  distinguishes them: only the max-responses validation constraint differs (`1` vs
  unset vs any larger number). The distinction does matter downstream, though — for
  stored-data formats, quotas, and especially SPSS conversion;
- **"info" items** — carry no response; they only carry metadata (labels, properties,
  …) with nothing stored.

**Non-scalar — blocks of fields / loops / grids:**

- "Loop" and "grid" are effectively the same thing (a different internal interface /
  type constant).
- They always carry a **block of fields** ("columns") repeated for every iteration.
  There is no array-of-int: to store an integer, a loop/grid adds a column of type
  `long`.
- A single-column grid/loop is conventionally named `GV` (not required by the format);
  numeric grids use `Num`, ranking questions use `Rank`.
- Columns can themselves be any structure (nested blocks/loops/grids/info items/
  scalars); deeper nesting is allowed, but in practice nobody needs more than 2–3
  levels.
- **Iterations** are either a list of categories (most common) or a numeric range
  (rare, but supported).
- The root MDD document is itself a block of fields, so the whole document is a tree.

## Categories

Categories are used as loop/grid iterations and in categorical questions. They are
first-class objects with a name, a label (context/language-aware), properties, and
modifiers.

Categories are not always real response options or iterations: some only group other
categories, like folders in a file system — they carry no data and need no database
column.

## Shared lists

Instead of inline categories, an object can reference a **shared list** of categories
defined once, globally, in the MDD document. A shared list is reachable in the object
model through `.Types` (see code).

There is no MDD syntax symbol for a shared-list reference; the `@` seen in the JSON is
only this toolchain's convention for recording a reference (see the reader's row-name
conventions in `docs/architecture.md`).

# Example outputs of `read_mdd`

These files are real outputs of the MDD reader (`python src/read_mdd.py --mdd ...`),
captured from small synthetic MDD projects — **not real/sensitive data**.

- `a1_v1.mdd.json`, `a1_v2.mdd.json` — two versions of a small survey ("a1").
- `b1_today.mdd.json`, `b1_tomorrow.mdd.json` — two versions of a second survey ("b1").

They are committed here as *example outputs* and used by `tests/test_fixtures.py` to
pin the JSON report shape. They cannot be regenerated in CI (reading an MDD requires
Windows + IBM/Unicom Professional), so they are the source of truth for what the
reader emits.

Downstream consumers (the diff tool, `mdmreportpy`, `mdmexcelreportpy`) accept exactly
this shape — the master definition lives in `mdmreportpy`'s
`docs/report-json-format.md` and `docs/rich-text-format.md`.

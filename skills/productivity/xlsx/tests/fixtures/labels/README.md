# Oracle labels (AC-B03)

The role-accuracy criterion is only meaningful if the ground truth is
NOT produced by the same logic it grades. Ideally an independent HUMAN
labels the fixtures.

Current state (documented honestly): on 2026-09-29 the user delegated
the labeling to the agent, so `labels.json` was authored by the agent
from the fixture-design rubric (role = what the column MEANS per the
dictionary; formula columns keep their business role). Measured:
58/61 = 95.1% (assertion >= 90%). This is NOT an independent human
oracle; replacing `labels.json` with your own labels re-runs the same
metric without any code change.

How to use:

1. Copy `labels.template.json` to `labels.json`.
2. Fill `expected_role` for each row with one of the dictionary roles
   (see `references/semantic-dictionary.tr-en.json` -> `roles`):
   identifier, label, date, period, quantity, currency, percent, code,
   address, person_name, free_text, summary, flag, formula_derived,
   empty_slot, unknown.
3. Run the oracle test:

       python -m pytest tests/test_xlsx_semantics.py -k labels -q

While `labels.json` is absent (or empty) the test is SKIPPED and the
final report states AC-B03 as NOT_TESTED -- an unmeasured accuracy is
never reported as a success.

Fixture variants (built by `make_book()` in the test file): clean,
english, reordered, mismatch, extra, missing, duplicate, hidden, lookup,
ambiguous, merged. `book` must be one of these names; the test builds
each book with its OWN variant (unknown names fall back to `clean`).

Labeling aid: `labels.worksheet.md` (next to this README) lists every
fixture column with its header, data type and sample values, plus the
machine's current guess - use it as a checklist, but the `expected_role`
values in `labels.json` remain YOUR decision (the agent never writes
`labels.json`; otherwise the oracle would grade itself).

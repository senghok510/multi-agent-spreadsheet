# First end-to-end run (Milestone 1)

One SpreadsheetBench V1 task was run through the full Inspector → Orchestrator → Editor → Validator
graph with a real model. This records what happened; it is not a benchmark score.

| | |
|---|---|
| Command | `uv run spreadsheet-agents --task-id 59196` |
| Task | `59196`, variant 1 (`sample_data_200`) |
| Instruction | "I need a formula to determine which column contains the highest value in a row, and then return the heading of that column." |
| Model | `gpt-4.1-mini`, temperature 0 |
| Code version | commit `1cdb7ad` |
| Latency | 11.5 s |
| Result | **failed** (runtime validation rejected the output) |
| Benchmark evaluated | **no**: reference answers were never loaded; `benchmark_evaluated` is `false` |

`status` is the runtime Validator's verdict. The official evaluation against the reference answer
is Milestone 3, so this run says nothing yet about task accuracy.

## What the agents did

1. **Inspector** found the headers in `D2:G2`, the data in `D3:G5`, and that column `H` is the output.
2. **Orchestrator** planned one formula, filled down.
3. **Editor** produced three operations: `set_formula` on `H3`
   (`=INDEX(D$2:G$2, MATCH(MAX(D3:G3), D3:G3, 0))`), then `copy_range` `H3` → `H4` and `H3` → `H5`.
4. **Validator** rejected the workbook: `H4` and `H5` hold `H3`'s formula unchanged, so they still
   compute row 3. It proposed per-row formulas such as
   `=INDEX(D$2:G$2, MATCH(MAX(D4:G4), D4:G4, 0))`.

## Finding

`copy_range` copies cell contents verbatim and does not adjust relative references the way Excel's
fill-down does. Any task that fills a formula down a column will produce wrong results through
`copy_range`. The runtime Validator caught it, but only because its LLM pass inspected the output;
the deterministic pass reports `copy_range` as unverified (a warning), as designed.

Not fixed in this milestone on purpose: changing the tool or prompts to pass a single task would
bias the baseline. Follow-up: make `copy_range` translate relative references (for example with
`openpyxl.formula.translate.Translator`) and verify copies deterministically.

## Problems the real run exposed (already fixed)

Mocked tests cannot find these; only calling the real API can.

- OpenAI's strict structured-output mode rejected the schemas (a `dict` field, defaulted fields).
  All agents now use `method="function_calling"`.
- The Editor wrote a `set_formula` value without the leading `=`; the Editor prompt now states the
  format. The tool's refusal was correct.
- A crash in one agent discarded the earlier agents' output; the runner now streams state so the
  trace keeps what finished.

## Caveats

- One task, one run. Model output varies between runs even at temperature 0.
- Token usage and cost are not recorded yet.
- The Validator does not check that unrelated cells were preserved: it never receives the original
  workbook.

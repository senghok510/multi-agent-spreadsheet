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

This was a tool bug, not a prompt problem, so it was fixed afterwards (issue #22): `copy_range` now
shifts relative references like Excel's fill-down, and the Validator checks copies against the
output. Prompts were not changed to make this task pass, to keep the baseline unbiased.

## After the `copy_range` fix: six tasks

Run on code version `fc427be` (`copy_range` fix plus the array-formula fix), `gpt-4.1-mini`, one run
each. "Our verdict" is the runtime Validator; "Checked by hand" compares the output with the input
workbook. Neither is a benchmark score.

| Task | Our verdict | Checked by hand |
|---|---|---|
| `59196` | success | Per-row formulas in `H4:H5`, `H3` untouched. Looks correct |
| `15027` | success | *Hand-check was wrong; see the correction below.* The model rewrote `B6:H6` with a different formula, which I judged wrong by comparing with the original formulas. The benchmark scores only `E1`, and its reference value (1019.4) equals what the model's formulas produce |
| `493-18` | failed | Correctly failed. The Editor cleared rows 2-5 (Ford, VW, Audi, Ferrari); the rule (name not in column F) says rows 2, 3 and 7 (Ford, VW, Fiat). Audi and Ferrari were erased wrongly. The Validator's stated reason was inaccurate |
| `55860` | error | The Editor declined ("ambiguous"). The sheet has 43 rows; the model saw 5 |
| `61-1` | error | The Editor declined (instruction says VBA macro). The sheet has 9 rows; the model saw 5 |
| `275-49` | error | The Editor declined (ambiguous). The sheet has 11 rows; the model saw 5 |

Of six tasks, one output looked correct by hand (see the correction below for `15027`). Findings:

1. **The model sees only 5 rows per sheet** (`max_preview_rows`). Three of the refusals involve
   sheets larger than that. Not yet shown to be the cause.
2. **A refusal is recorded as `status: error`**, indistinguishable from a crash. It needs its own
   status (for example `declined`) so metrics can separate model refusals from code failures.
3. **"Success" means the plan was applied, not that it is right.** Nothing evaluates formulas, so
   a wrong formula can pass, and (as `15027` showed) a hand-check can be wrong too. The official
   evaluation against the reference answers is the only reliable judge; see the scoring section.
4. **The Validator's explanations are unreliable.** It filed a passing observation under errors
   (`15027`, earlier run) and gave the wrong reason for `493-18`.
5. The Editor adds changes nobody asked for (`493-18` wrote Keep/Remove formulas into column G).

## After showing the model whole sheets (50 rows x 30 columns)

Finding 1 above suspected the 5-row preview. The default view is now 50 rows by 30 columns
(79% of the dataset's sheets in full; sheets that do not fit are flagged `truncated`). The three
declined tasks were rerun, one run each, `gpt-4.1-mini`:

| Task | Before | Now | Checked by hand |
|---|---|---|---|
| `55860` | declined | **declined** | All 43 rows were visible and the Editor still called the request ambiguous, so the preview was not the cause here |
| `61-1` | declined | success | **False success.** Column D is unchanged; the Editor wrote `=D2 & "-001"` into a new column G. The task changes D in place. (The reference answer shows `-001` appended three times, so a single append would not match it either.) The Validator noted the workaround and still passed it |
| `275-49` | declined | failed | Correct verdict, wrong reason. The Validator complained that no macro exists; it missed that the Editor wrote `=B2` into `B2` and `=C2` into `C2` (circular references that destroy the original data). Column A, an ID, was overwritten with a merged string |

Conclusion: seeing the whole sheet is the right input, but it did not make results more correct.
Two refusals became attempts that did not pass the benchmark, one of which passed our checks.

New findings:

6. **Self-referencing formulas are not rejected.** `set_formula` accepted `=B2` in cell `B2`. This is
   objectively wrong and cheap to check deterministically.
7. **The Editor tends to add helper columns instead of editing in place**, and the Validator accepts
   the workaround even when it notes it.

## Scored against the reference answers

`spreadsheet_agents.benchmark.evaluator` (issue #30) recalculates each output with headless
LibreOffice and compares the cell values in the task's `answer_position` with the reference
answer, following the benchmark's own `evaluation.py`: values only, numbers rounded to 2 decimals,
exact strings, empty equals empty. It is separate from the agents and never imported by them.

Newest run of each task, **variant 1 only**:

| Task | Our runtime verdict | Benchmark | Detail |
|---|---|---|---|
| `59196` | success | **pass** | |
| `15027` | success | **pass** | `E1` = 1019.4 as in the reference |
| `55860` | error (declined) | fail | no output workbook |
| `61-1` | success | **fail** | `D2` expected `RPC0094-001-001-001`, got `RPC0094`: column D was never changed |
| `493-18` | failed | fail | `A2` expected `Audi`, got empty: Audi, Ferrari and Porsche were erased |
| `275-49` | failed | fail | `#VALUE!` from the circular references written before the self-reference check existed |

Two of six pass. Our runtime Validator agreed with the benchmark on five of six; its one error was
`61-1`, a false success.

**Correction.** An earlier version of this note called `15027` a false success, because the model's
formulas in `B6:H6` differ from the original ones. That was a wrong inference: the task is scored on
`E1` only, and the reference value matches the model's result. A hand-check against an assumed
target is not a substitute for the benchmark's own definition of correct.

Limits of these numbers: one run per task, variant 1 only (the benchmark also runs variants 2 and 3,
and its hard score needs all three), and six of 200 tasks.

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

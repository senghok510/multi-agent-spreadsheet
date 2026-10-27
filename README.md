# Multi-Agent Spreadsheet Automation

Build a reproducible multi-agent AI system that understands natural-language spreadsheet tasks, inspects Excel workbooks, plans and executes modifications, validates results, and evaluates performance on SpreadsheetBench. Develop the system **incrementally: one feature branch, one focused pull request, and one merge into `main` at a time**.

## 1) Project Goal

- Automate Excel spreadsheet tasks from natural-language instructions.
- Build specialized agents for **inspection, planning/orchestration, editing, validation**, and later debugging and visualization.
- Use **LangGraph** to manage shared state, task dependencies, conditional routing, and retries.
- Start with a reproducible **baseline** before introducing prompt/skill optimization.
- Evaluate task success, workbook preservation, latency, token usage, and cost on **SpreadsheetBench V1**, then expand to **SpreadsheetBench 2**.
- Compare a single-agent baseline with a multi-agent system and, later, a SkillOpt-enhanced variant.

## 2) End-to-End Architecture (High Level)

1. **Task Input & Benchmark Adapter**
   - Accept a user instruction and input `.xlsx` workbook.
   - Load benchmark task metadata and workbook variants.
   - Keep reference answers and evaluation metadata **outside the agent runtime**.

2. **Inspection Agent**
   - Read sheet names, dimensions, headers, representative cells, and formulas.
   - Identify relevant ranges and potential ambiguities.
   - Return structured workbook context without modifying the file.

3. **Orchestrator (Planning + Delegation)**
   - Break the request into dependent subtasks.
   - Decide which specialized agent handles each subtask.
   - Track task progress, workbook versions, and execution state.
   - **Milestone 1:** use a fixed Inspector → Orchestrator → Editor → Validator graph.
   - **Later:** evolve toward Anthropic's **Orchestrator–Workers** pattern with dynamic delegation.

4. **Editing Agent & Workbook Tools**
   - Translate planned steps into structured, validated spreadsheet operations.
   - Execute operations with `openpyxl` and other deterministic tools.
   - Write results to a **new workbook copy**, preserving the input.
   - Do not execute arbitrary LLM-generated Python in the host process.

5. **Validation Agent**
   - Check workbook readability, expected edits, and preservation of unrelated content.
   - Report failures and uncertainty; do not claim correctness merely because a file was saved.
   - **Later:** add an Evaluator–Optimizer loop: validate → debug/revise → revalidate.

6. **Visualization Agent (Later Milestone)**
   - Create charts and presentation-oriented formatting when tasks require them.

7. **Debugging Agent (Later Milestone)**
   - Diagnose tool errors and failed validations.
   - Propose bounded corrections and retries.

8. **External Benchmark Evaluator**
   - Compare predictions against reference answers using the benchmark's evaluation protocol.
   - Keep the evaluator separate from the runtime Validation Agent.
   - Record success/failure, latency, token usage, and costs per task.

9. **Observability & Reproducibility**
   - Save structured agent traces, task plans, operations, validation reports, and metrics.
   - Record model, prompt version, configuration, dataset split, and code commit for each experiment.

## 3) Core Design Principles

- **No answer leakage:** agents receive the instruction and input workbook, never the reference answer.
- **Preserve originals:** execute on isolated copies and avoid overwriting source workbooks.
- **Agents decide; tools execute:** LLMs plan and choose operations; deterministic, validated tools perform edits.
- **Structured interfaces:** use Pydantic schemas for agent outputs and LangGraph state.
- **Bounded execution:** restrict operation types, file access, execution time, and retries; sandbox generated code if introduced later.
- **Honest validation:** distinguish runtime checks from benchmark ground-truth evaluation.
- **Reproducible baselines:** fix prompts and evaluation settings before comparing methods.
- **Incremental delivery:** each feature should be implemented, tested, reviewed in a PR, and merged before starting the next dependent feature.

## 4) Data & Evaluation Lifecycle

### Phase A — Dataset Research
- Download and inspect SpreadsheetBench V1 (`sample_data_200` for development).
- Identify task instructions, workbook variants, answer files, and evaluation metadata.
- Explore representative tasks and document supported vs. unsupported operations.
- Split benchmark tasks without allowing workbook variants of the same task to cross splits; check for overlap between dataset subsets.

### Phase B — Baseline Workflow
- Implement workbook inspection and a limited operation interpreter.
- Add structured orchestrator planning and editor operation generation.
- Connect Inspector → Orchestrator → Editor → Validator with LangGraph.
- Run one simple synthetic task, then a real benchmark task.

### Phase C — Agent Specialization
- Add dynamic delegation to editing, visualization, and debugging workers.
- Add validation-triggered retries with a maximum retry budget.
- Introduce stronger workbook-preservation checks.

### Phase D — Benchmark Experiments
- Evaluate a single-agent baseline and multi-agent baseline on the same held-out tasks.
- Measure task success, preservation, latency, tokens, and cost.
- Run error analysis and ablations to identify which agents add value.

### Phase E — Skill Optimization
- Integrate Microsoft SkillOpt **after** baseline measurements are stable.
- Optimize agent instructions/skills using development data only.
- Evaluate on held-out tasks without tuning against test answers.

## 5) Suggested Repository Layout

```text
multi-agent-spreadsheet/
├── README.md
├── pyproject.toml
├── .env.example
├── .gitignore
├── configs/
│   ├── baseline.yaml
│   └── experiments/
├── data/
│   ├── raw/                  # local datasets; not committed
│   └── fixtures/             # tiny synthetic workbooks for tests
├── src/
│   └── spreadsheet_agents/
│       ├── agents/
│       │   ├── inspector.py
│       │   ├── orchestrator.py
│       │   ├── editor.py
│       │   ├── validator.py
│       │   ├── debugger.py    # later
│       │   └── visualizer.py  # later
│       ├── graph/
│       │   ├── state.py
│       │   ├── builder.py
│       │   └── routing.py
│       ├── prompts/
│       │   └── baseline/
│       ├── tools/
│       │   ├── workbook_reader.py
│       │   ├── workbook_editor.py
│       │   └── code_executor.py
│       ├── schemas/
│       │   ├── task.py
│       │   └── workbook.py
│       ├── benchmark/
│       │   ├── spreadsheetbench_v1.py
│       │   └── evaluator.py   # later
│       ├── optimization/      # later: SkillOpt integration
│       └── cli.py
├── experiments/
│   └── run_baseline.py
├── tests/
│   ├── unit/
│   └── integration/
├── outputs/                   # generated workbooks/traces/metrics; not committed
├── docs/
│   └── architecture.md
└── .github/
    └── workflows/
        └── ci.yml
```

**Note:** This is the target layout. Create files as their features are implemented; do not fill every module with placeholder logic on day one.

## 6) Incremental Development Workflow (Branch → PR → Main)

Use a short-lived branch for each **small, independently testable feature**. Avoid implementing the entire pipeline in one giant branch.

1. Start from an up-to-date `main` branch.
2. Create a focused branch: `feat/workbook-inspection`, `feat/graph-state`, `test/editor-operations`, etc.
3. Implement one coherent change with tests and documentation.
4. Run local checks (`pytest`, `ruff`) and manually verify behavior where appropriate.
5. Commit with a descriptive message and push the branch.
6. Open a pull request describing **what changed, why, how it was tested, and known limitations**.
7. Review the diff, address feedback, and merge only after CI passes.
8. Delete the merged branch and start the next feature from the latest `main`.

Example commands:

```bash
git switch main
git pull --ff-only
git switch -c feat/workbook-inspection

# implement the feature and its tests
uv run pytest
uv run ruff check .

git add .
git commit -m "feat: add workbook inspection tool"
git push -u origin feat/workbook-inspection

# Open a PR on GitHub; merge after review and passing checks.
```

**Branch naming:** `feat/`, `fix/`, `test/`, `docs/`, `refactor/`.

**Definition of Done for a feature:** working code, tests, documented behavior, no exposed secrets or benchmark answers, passing CI, and a merged PR.

## 7) Execution Roadmap (Small PRs)

### Milestone 1 — Working Baseline

1. **PR 01 — Repository foundation:** `pyproject.toml`, dependency management, `.gitignore`, README, pytest/Ruff, CI.
2. **PR 02 — Workbook inspection:** inspect sheets, dimensions, headers, sample rows, and formulas; add unit tests.
3. **PR 03 — Structured contracts:** Pydantic task/operation schemas and LangGraph shared state.
4. **PR 04 — Workbook editing tools:** restricted Excel operations, output-copy handling, and tests.
5. **PR 05 — Inspection + planning agents:** baseline prompts, LLM integration, and structured task plans.
6. **PR 06 — Editing agent:** convert plans into validated operations and execute them.
7. **PR 07 — Validation agent:** deterministic checks, preservation checks, and structured reports.
8. **PR 08 — LangGraph integration:** connect the four agents and test the full workflow with mocked LLM responses.
9. **PR 09 — SpreadsheetBench adapter:** load a real task instruction and input workbook without exposing answer files to agents.
10. **PR 10 — CLI + first experiment:** run one task end-to-end and save workbook, trace, and basic metrics.

### Milestone 2 — Robust Multi-Agent Orchestration

11. Add dynamic Orchestrator–Workers delegation and dependency-aware routing.
12. Add Debugging Agent and bounded validation/retry loop.
13. Add Visualization Agent and chart/formatting tools.
14. Improve execution isolation, workbook fidelity, and observability.

### Milestone 3 — Reproducible Benchmarking

15. Implement the official benchmark evaluation protocol.
16. Run single-agent vs. multi-agent comparisons on held-out tasks.
17. Add ablations, failure categorization, latency/cost analysis, and reports.

### Milestone 4 — SkillOpt Research

18. Add SkillOpt adapter and optimized skill configurations.
19. Optimize on development tasks and evaluate on untouched test tasks.
20. Compare baseline vs. optimized system and publish reproducible findings.

## 8) Minimal Success Criteria (Milestone 1)

- A local command accepts an instruction and `.xlsx` input workbook.
- The Inspector returns structured workbook information.
- The Orchestrator produces a structured plan.
- The Editor executes at least a small, documented set of operations.
- The Validator detects basic mismatches and reports limitations.
- LangGraph runs the four-agent workflow end-to-end.
- A new output workbook and execution trace are saved.
- At least one real SpreadsheetBench task is attempted and its outcome is recorded accurately.
- Tests and lint checks pass in GitHub Actions.
- Every milestone feature is delivered through a focused branch and pull request.

## 9) Key Risks to Watch Early

- **Benchmark leakage:** reference answers accidentally appearing in agent prompts or runtime state.
- **Workbook damage:** loss of formulas, styles, charts, merged cells, or unsupported Excel features.
- **False validation success:** operations appear applied but the user's task remains incomplete.
- **Unsafe code execution:** generated Python or untrusted workbooks accessing host resources.
- **Dataset/schema assumptions:** metadata and evaluation rules differ between releases.
- **LLM variability:** changes in model versions, prompts, or sampling settings affect comparisons.
- **Agent overhead:** multi-agent workflows may cost more and take longer without improving accuracy.
- **Oversized PRs:** mixing many features makes debugging, review, and rollback difficult.

## 10) Initial Technology Stack

| Area | Technology |
|---|---|
| Language & environment | Python 3.11+, `uv` |
| Agent orchestration | LangGraph |
| LLM integration | OpenAI API / LangChain OpenAI |
| Structured outputs | Pydantic |
| Excel operations | openpyxl, pandas where appropriate |
| Testing & quality | pytest, Ruff, GitHub Actions |
| Dataset | SpreadsheetBench V1; SpreadsheetBench 2 later |
| Experiment outputs | JSON traces, metrics, generated `.xlsx` files |
| Future skill optimization | Microsoft SkillOpt |

## 11) References

- [Anthropic — Building Effective Agents](https://www.anthropic.com/engineering/building-effective-agents)
- [LangGraph documentation](https://docs.langchain.com/oss/python/langgraph/overview)
- [SpreadsheetBench V1](https://github.com/RUCKBReasoning/SpreadsheetBench)
- [SpreadsheetBench 2](https://github.com/RUCKBReasoning/SpreadsheetBench-2)
- [Microsoft SkillOpt](https://github.com/microsoft/SkillOpt)

---

**Development rule:** Keep `main` working. Build one feature at a time, validate it locally, open a focused PR, and merge only when it is tested. First prove a minimal end-to-end baseline; then add complexity only when experiments show it is useful.

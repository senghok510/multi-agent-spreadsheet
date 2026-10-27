# Repository Guidelines

## Project Structure & Module Organization

Python source lives in `src/spreadsheet_agents/`. Keep agent behavior in `agents/`, workflow construction and routing in `graph/`, workbook operations in `tools/`, Pydantic models in `schemas/`, and dataset integration in `benchmark/`. Baseline prompts live in `prompts/baseline/` inside the package.

Use `configs/baseline.yaml` for experiment settings and `experiments/run_baseline.py` for the future baseline runner. Tests live in `tests/`. Benchmark inputs are under `data/raw/spreadsheetbench_v1/sample_data_200/`; generated workbooks, traces, and metrics belong in `outputs/`. Most modules currently contain placeholders.

## Build, Test, and Development Commands

Run commands from the repository root with Python 3.11 or newer:

- `uv sync --extra dev`: install the project and development dependencies from the lockfile.
- `uv run spreadsheet-agents --help`: inspect the scaffold CLI.
- `uv run pytest`: run the test suite.
- `uv run ruff check .`: check Python style and common errors.
- `uv run ruff format .`: format Python files.
- `uv build`: build source and wheel distributions using Hatchling.

The CLI does not execute agents yet; the baseline runner currently exits with an implementation-pending message. Update `pyproject.toml` and `uv.lock` together when changing dependencies.

## Coding Style & Naming Conventions

Use four-space indentation, type annotations on public interfaces, and concise module docstrings. Use `snake_case` for modules, functions, and variables; `PascalCase` for classes; and `UPPER_SNAKE_CASE` for constants. Keep workbook I/O in tools and workflow transitions in graph modules. Ruff is available in the development extra; use its defaults until project-specific settings are added.

## Testing Guidelines

Use pytest with files named `test_*.py` and functions named `test_<behavior>`. Create small workbooks in temporary directories for editing and validation tests. Mock model responses to keep routine tests deterministic and independent of API credentials. Cover successful edits, invalid inputs, and workflow termination as functionality is implemented. No coverage threshold is configured; current test files are placeholders and collect no tests.

## Commit & Pull Request Guidelines

There are no commits yet, so no historical message convention exists. Use concise imperative subjects, such as `Add workbook inspection schema`. PRs should explain the change, link an issue when available, and report validation results and limitations.

## Security & Configuration

Store credentials in `.env`, using `.env.example` as a template. Keep raw inputs unchanged and save edits under `outputs/workbooks/`. Restrict benchmark answer workbooks to evaluation; exclude them from agent context. Raw data, generated artifacts, and secrets are ignored by Git.

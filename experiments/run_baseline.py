"""Run the baseline on one SpreadsheetBench task, using configs/baseline.yaml.

Usage: uv run python experiments/run_baseline.py <task_id>
"""

import sys

from spreadsheet_agents.cli import main

if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    raise SystemExit(main(["--task-id", sys.argv[1]]))

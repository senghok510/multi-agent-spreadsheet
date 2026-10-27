"""Run the four-agent workflow on one task and record what happened."""

import json
import subprocess
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from spreadsheet_agents.graph.builder import build_graph


def _git_commit() -> str | None:
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True,
            text=True,
            check=True,
        )
    except (OSError, subprocess.CalledProcessError):
        return None

    return result.stdout.strip() or None


def _write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")


def run_task(
    *,
    instruction: str,
    input_path: str | Path,
    model: str,
    output_dir: str | Path = "outputs",
    label: str | None = None,
    task_id: str | None = None,
    variant: int | None = None,
) -> dict[str, Any]:
    """Run one task and write a workbook, trace and metrics file.

    The returned metrics describe the runtime validator's verdict only.
    They are not a benchmark score: reference answers are never loaded here.
    """

    input_path = Path(input_path)
    output_dir = Path(output_dir)

    started = datetime.now(UTC)
    run_id = f"{started:%Y%m%dT%H%M%SZ}_{label or input_path.stem}"

    workbook_path = output_dir / "workbooks" / f"{run_id}.xlsx"
    trace_path = output_dir / "traces" / f"{run_id}.json"
    metrics_path = output_dir / "metrics" / f"{run_id}.json"

    metrics: dict[str, Any] = {
        "run_id": run_id,
        "task_id": task_id,
        "variant": variant,
        "model": model,
        "git_commit": _git_commit(),
        "started_at": started.isoformat(),
        "input_path": str(input_path),
        "benchmark_evaluated": False,
    }

    final: dict[str, Any] = {}
    clock = time.perf_counter()

    try:
        final = build_graph().invoke(
            {
                "user_request": instruction,
                "input_path": str(input_path),
                "output_path": str(workbook_path),
                "model": model,
                "trace": [],
            }
        )
        metrics["status"] = final["status"]
        metrics["error"] = final.get("error")
    except Exception as exc:
        metrics["status"] = "error"
        metrics["error"] = f"{type(exc).__name__}: {exc}"

    metrics["latency_seconds"] = round(time.perf_counter() - clock, 3)

    edit_plan = final.get("edit_plan") or {}
    metrics["operations_planned"] = len(edit_plan.get("operations", []))
    metrics["runtime_validation"] = final.get("validation_report")
    metrics["output_path"] = str(workbook_path) if workbook_path.exists() else None
    metrics["trace_path"] = str(trace_path)
    metrics["metrics_path"] = str(metrics_path)

    _write_json(
        trace_path,
        {
            "run_id": run_id,
            "instruction": instruction,
            "steps": final.get("trace", []),
            "inspection_report": final.get("inspection_report"),
            "task_plan": final.get("task_plan"),
            "edit_plan": final.get("edit_plan"),
            "validation_report": final.get("validation_report"),
            "error": metrics["error"],
        },
    )
    _write_json(metrics_path, metrics)

    return metrics

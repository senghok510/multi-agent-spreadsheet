"""Command-line entry point: run the multi-agent workflow on one task."""

import argparse
import os
from pathlib import Path

import yaml
from dotenv import load_dotenv

from spreadsheet_agents.benchmark.spreadsheetbench_v1 import load_task
from spreadsheet_agents.runner import run_task

DEFAULT_CONFIG = "configs/baseline.yaml"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the multi-agent spreadsheet workflow on one task."
    )
    parser.add_argument("--instruction", help="Natural-language task (with --input)")
    parser.add_argument("--input", help="Path to the input .xlsx workbook")
    parser.add_argument("--task-id", help="SpreadsheetBench task id (instead of the two above)")
    parser.add_argument("--variant", type=int, help="Benchmark workbook variant (default: config)")
    parser.add_argument("--benchmark-root", help="Benchmark folder (default: config)")
    parser.add_argument("--model", help="Model name (default: config)")
    parser.add_argument("--output-dir", help="Where to write outputs (default: config)")
    parser.add_argument("--config", default=DEFAULT_CONFIG, help="YAML config path")
    return parser


def load_config(path: str | Path) -> dict:
    with Path(path).open(encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def main(argv: list[str] | None = None) -> int:
    load_dotenv()
    parser = build_parser()
    args = parser.parse_args(argv)

    has_task_id = args.task_id is not None
    has_custom = args.instruction is not None or args.input is not None

    if has_task_id == has_custom:
        parser.error("give either --task-id, or both --instruction and --input")

    if has_custom and not (args.instruction and args.input):
        parser.error("--instruction and --input must be used together")

    if not os.environ.get("OPENAI_API_KEY"):
        parser.error("OPENAI_API_KEY is not set (put it in .env or the environment)")

    config = load_config(args.config)
    model = args.model or config.get("llm", {}).get("model")
    output_dir = args.output_dir or config.get("output", {}).get("directory", "outputs")

    if not model:
        parser.error("no model: pass --model or set llm.model in the config")

    if has_task_id:
        benchmark = config.get("benchmark", {})
        root = args.benchmark_root or benchmark.get("root")
        variant = args.variant or benchmark.get("variant", 1)

        if not root:
            parser.error("no benchmark root: pass --benchmark-root or set benchmark.root")

        task = load_task(root, args.task_id, variant)
        run = {
            "instruction": task.instruction,
            "input_path": task.input_path,
            "label": task.task_id,
            "task_id": task.task_id,
            "variant": task.variant,
        }
    else:
        run = {"instruction": args.instruction, "input_path": args.input}

    metrics = run_task(model=model, output_dir=output_dir, **run)

    print(f"status:      {metrics['status']}")
    if metrics["error"]:
        print(f"error:       {metrics['error']}")
    print(f"latency:     {metrics['latency_seconds']}s")
    print(f"workbook:    {metrics['output_path']}")
    print(f"trace:       {metrics['trace_path']}")
    print(f"metrics:     {metrics['metrics_path']}")
    print("note: status is the runtime validator's verdict, not a benchmark score")

    return 0 if metrics["status"] == "success" else 1


if __name__ == "__main__":
    raise SystemExit(main())

"""Tests for the command-line interface (the workflow itself is faked)."""

import json
from pathlib import Path

import pytest

from spreadsheet_agents import cli
from spreadsheet_agents.cli import load_config, main


@pytest.fixture(autouse=True)
def api_key(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "test-key")


@pytest.fixture
def config(tmp_path: Path) -> Path:
    path = tmp_path / "config.yaml"
    path.write_text(
        "llm:\n  model: config-model\n"
        "benchmark:\n  root: BENCH\n  variant: 1\n"
        "output:\n  directory: out-from-config\n"
    )
    return path


@pytest.fixture
def calls(monkeypatch):
    """Replace run_task; record its arguments and return a canned result."""
    recorded: dict = {}
    result = {
        "status": "success",
        "error": None,
        "latency_seconds": 1.5,
        "output_path": "o.xlsx",
        "trace_path": "t.json",
        "metrics_path": "m.json",
    }

    def fake_run_task(**kwargs):
        recorded.update(kwargs)
        return result

    monkeypatch.setattr(cli, "run_task", fake_run_task)
    recorded["result"] = result
    return recorded


def test_custom_task_passes_instruction_and_input(calls, config):
    code = main(["--instruction", "do it", "--input", "a.xlsx", "--config", str(config)])

    assert code == 0
    assert calls["instruction"] == "do it"
    assert calls["input_path"] == "a.xlsx"
    assert calls["model"] == "config-model"
    assert calls["output_dir"] == "out-from-config"


def test_flags_override_the_config(calls, config):
    main(
        [
            "--instruction", "x", "--input", "a.xlsx", "--config", str(config),
            "--model", "other-model", "--output-dir", "elsewhere",
        ]
    )

    assert calls["model"] == "other-model"
    assert calls["output_dir"] == "elsewhere"


def test_task_id_loads_instruction_and_input_from_the_benchmark(calls, config, tmp_path):
    root = tmp_path / "bench"
    (root / "spreadsheet" / "101").mkdir(parents=True)
    (root / "dataset.json").write_text(json.dumps([{"id": 101, "instruction": "Sum A"}]))
    (root / "spreadsheet" / "101" / "1_101_input.xlsx").write_bytes(b"")
    (root / "spreadsheet" / "101" / "1_101_answer.xlsx").write_bytes(b"")

    code = main(["--task-id", "101", "--benchmark-root", str(root), "--config", str(config)])

    assert code == 0
    assert calls["instruction"] == "Sum A"
    assert calls["input_path"].name == "1_101_input.xlsx"
    assert calls["task_id"] == "101"
    assert calls["variant"] == 1
    assert "answer_path" not in calls
    assert "answer" not in str(calls["input_path"]).replace("input", "")


def test_nonzero_exit_code_when_run_does_not_succeed(calls, config):
    calls["result"]["status"] = "failed"
    calls["result"]["error"] = "bad"

    assert main(["--instruction", "x", "--input", "a.xlsx", "--config", str(config)]) == 1


def test_output_says_status_is_not_a_benchmark_score(calls, config, capsys):
    main(["--instruction", "x", "--input", "a.xlsx", "--config", str(config)])
    assert "not a benchmark score" in capsys.readouterr().out


@pytest.mark.parametrize(
    "argv",
    [
        [],
        ["--instruction", "x"],
        ["--input", "a.xlsx"],
        ["--task-id", "1", "--instruction", "x", "--input", "a.xlsx"],
    ],
)
def test_invalid_argument_combinations_exit_with_usage_error(argv, config, calls):
    with pytest.raises(SystemExit) as exc:
        main([*argv, "--config", str(config)])
    assert exc.value.code == 2


def test_missing_api_key_is_a_clear_error(monkeypatch, config, calls, capsys):
    monkeypatch.delenv("OPENAI_API_KEY")
    monkeypatch.setattr(cli, "load_dotenv", lambda: None)

    with pytest.raises(SystemExit) as exc:
        main(["--instruction", "x", "--input", "a.xlsx", "--config", str(config)])

    assert exc.value.code == 2
    assert "OPENAI_API_KEY" in capsys.readouterr().err


def test_repo_config_has_the_settings_the_cli_reads():
    config = load_config("configs/baseline.yaml")
    assert config["llm"]["model"]
    assert config["benchmark"]["root"]
    assert config["output"]["directory"]

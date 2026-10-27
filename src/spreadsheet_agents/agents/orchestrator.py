"""Plan tasks and coordinate the spreadsheet agents.
"""
from pathlib import Path

from langchain_openai import ChatOpenAI

from spreadsheet_agents.agents.inspector import InspectionReport
from spreadsheet_agents.schemas.task import TaskPlan
from spreadsheet_agents.schemas.workbook import WorkbookInfo

PROMPT_PATH = (
    Path(__file__).resolve().parents[1]
    / "prompts/baseline/orchestrator.md"
)


def run_orchestrator(
    user_request: str,
    workbook_info: WorkbookInfo,
    inspection_report: InspectionReport,
    model: str,
) -> TaskPlan:

    llm = ChatOpenAI(
        model=model,
        temperature=0,
    ).with_structured_output(TaskPlan)

    prompt = PROMPT_PATH.read_text(encoding="utf-8")

    plan = llm.invoke([
        ("system", prompt),
        (
            "human",
            f"User request:\n{user_request}\n\n"
            f"Workbook:\n"
            f"{workbook_info.model_dump_json(indent=2)}\n\n"
            f"Inspection report:\n"
            f"{inspection_report.model_dump_json(indent=2)}"
        ),
    ])

    return plan
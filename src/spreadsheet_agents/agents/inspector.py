"""Inspect workbook structure and identify task-relevant cells.

"""
from pathlib import Path
from langchain_openai import ChatOpenAI
from pydantic import BaseModel, Field

from spreadsheet_agents.schemas.workbook import WorkbookInfo
from spreadsheet_agents.tools.workbook_reader import inspect_workbook


PROMPT_PATH = (
    Path(__file__).resolve().parents[1]
    / "prompts/baseline/inspector.md"
)


class InspectionReport(BaseModel):
    relevant_sheets: list[str] = Field(default_factory=list)
    relevant_columns: dict[str, list[str]] = Field(
        default_factory=dict
    )
    observations: list[str] = Field(default_factory=list)
    ambiguities: list[str] = Field(default_factory=list)


def run_inspector(
    workbook_path: str,
    user_request: str,
    model: str,
) -> tuple[WorkbookInfo, InspectionReport]:

    workbook_info = inspect_workbook(workbook_path)

    llm = ChatOpenAI(
        model=model,
        temperature=0,
    ).with_structured_output(InspectionReport)

    prompt = PROMPT_PATH.read_text(encoding="utf-8")

    report = llm.invoke([
        ("system", prompt),
        (
            "human",
            f"User request:\n{user_request}\n\n"
            f"Workbook schema:\n"
            f"{workbook_info.model_dump_json(indent=2)}"
        ),
    ])

    return workbook_info, report
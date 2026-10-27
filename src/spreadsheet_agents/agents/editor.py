"""Apply planned workbook changes to a separate output copy.
"""
from pathlib import Path

from langchain_openai import ChatOpenAI

from spreadsheet_agents.schemas.task import TaskPlan, EditPlan
from spreadsheet_agents.schemas.workbook import WorkbookInfo
from spreadsheet_agents.agents.inspector import InspectionReport
from spreadsheet_agents.tools.code_executor import execute_edit_plan


PROMPT_PATH = (
    Path(__file__).resolve().parents[1]
    / "prompts/baseline/editor.md"
)


def generate_edit_plan(
    user_request: str,
    task_plan: TaskPlan,
    workbook_info: WorkbookInfo,
    inspection_report: InspectionReport,
    model: str,
) -> EditPlan:

    llm = ChatOpenAI(
        model=model,
        temperature=0,
    ).with_structured_output(EditPlan)

    prompt = PROMPT_PATH.read_text(encoding="utf-8")

    edit_plan = llm.invoke([
        ("system", prompt),
        (
            "human",
            f"User request:\n{user_request}\n\n"
            f"Task plan:\n"
            f"{task_plan.model_dump_json(indent=2)}\n\n"
            f"Workbook schema:\n"
            f"{workbook_info.model_dump_json(indent=2)}\n\n"
            f"Inspection:\n"
            f"{inspection_report.model_dump_json(indent=2)}"
        ),
    ])

    if not edit_plan.operations:
        raise ValueError(
            "Editor generated no operations. "
            f"Explanation: {edit_plan.explanation}"
        )

    return edit_plan


def run_editor(
    user_request: str,
    task_plan: TaskPlan,
    workbook_info: WorkbookInfo,
    inspection_report: InspectionReport,
    input_path: str,
    output_path: str,
    model: str,
) -> EditPlan:

    edit_plan = generate_edit_plan(
        user_request=user_request,
        task_plan=task_plan,
        workbook_info=workbook_info,
        inspection_report=inspection_report,
        model=model,
    )

    execute_edit_plan(
        input_path=input_path,
        output_path=output_path,
        edit_plan=edit_plan,
    )

    return edit_plan

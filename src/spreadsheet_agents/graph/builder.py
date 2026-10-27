"""Construct the spreadsheet agent workflow.
"""
from datetime import UTC, datetime

from langgraph.graph import END, START, StateGraph

from spreadsheet_agents.agents.editor import run_editor
from spreadsheet_agents.agents.inspector import (
    InspectionReport,
    run_inspector,
)
from spreadsheet_agents.agents.orchestrator import run_orchestrator
from spreadsheet_agents.agents.validator import run_validator
from spreadsheet_agents.graph.routing import route_after_validation
from spreadsheet_agents.graph.state import SpreadsheetState
from spreadsheet_agents.schemas.task import (
    EditPlan,
    TaskPlan,
)
from spreadsheet_agents.schemas.workbook import WorkbookInfo


def make_trace(agent: str, status: str) -> dict:
    return {
        "agent": agent,
        "status": status,
        "timestamp": datetime.now(
            UTC
        ).isoformat(),
    }


def inspector_node(state: SpreadsheetState):

    info, report = run_inspector(
        workbook_path=state["input_path"],
        user_request=state["user_request"],
        model=state["model"],
    )

    return {
        "workbook_info": info.model_dump(),
        "inspection_report": report.model_dump(),
        "trace": state.get("trace", []) + [
            make_trace("inspector", "completed")
        ],
    }


def orchestrator_node(state: SpreadsheetState):

    plan = run_orchestrator(
        user_request=state["user_request"],
        workbook_info=WorkbookInfo.model_validate(
            state["workbook_info"]
        ),
        inspection_report=InspectionReport.model_validate(
            state["inspection_report"]
        ),
        model=state["model"],
    )

    return {
        "task_plan": plan.model_dump(),
        "trace": state.get("trace", []) + [
            make_trace("orchestrator", "completed")
        ],
    }


def editor_node(state: SpreadsheetState):

    edit_plan = run_editor(
        user_request=state["user_request"],
        task_plan=TaskPlan.model_validate(
            state["task_plan"]
        ),
        workbook_info=WorkbookInfo.model_validate(
            state["workbook_info"]
        ),
        inspection_report=InspectionReport.model_validate(
            state["inspection_report"]
        ),
        input_path=state["input_path"],
        output_path=state["output_path"],
        model=state["model"],
    )

    return {
        "edit_plan": edit_plan.model_dump(),
        "trace": state.get("trace", []) + [
            make_trace("editor", "completed")
        ],
    }


def validator_node(state: SpreadsheetState):

    report = run_validator(
        user_request=state["user_request"],
        task_plan=TaskPlan.model_validate(
            state["task_plan"]
        ),
        edit_plan=EditPlan.model_validate(
            state["edit_plan"]
        ),
        output_path=state["output_path"],
        model=state["model"],
    )

    return {
        "validation_report": report.model_dump(),
        "trace": state.get("trace", []) + [
            make_trace("validator", "completed")
        ],
    }


def success_node(state: SpreadsheetState):
    return {
        "status": "success",
        "error": None,
    }


def failed_node(state: SpreadsheetState):

    report = state.get("validation_report", {})

    return {
        "status": "failed",
        "error": "; ".join(report.get("errors", []))
        or "Validation failed",
    }


def build_graph():

    graph = StateGraph(SpreadsheetState)

    graph.add_node("inspector", inspector_node)
    graph.add_node("orchestrator", orchestrator_node)
    graph.add_node("editor", editor_node)
    graph.add_node("validator", validator_node)
    graph.add_node("success", success_node)
    graph.add_node("failed", failed_node)

    graph.add_edge(START, "inspector")
    graph.add_edge("inspector", "orchestrator")
    graph.add_edge("orchestrator", "editor")
    graph.add_edge("editor", "validator")

    graph.add_conditional_edges(
        "validator",
        route_after_validation,
        {
            END: "success",
            "failed": "failed",
        },
    )

    graph.add_edge("success", END)
    graph.add_edge("failed", END)

    return graph.compile()
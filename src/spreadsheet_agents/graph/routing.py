"""Define workflow transitions and termination rules.
"""
from langgraph.graph import END

from spreadsheet_agents.graph.state import SpreadsheetState


def route_after_validation(state: SpreadsheetState) -> str:

    report = state.get("validation_report", {})
    ## pass or not
    if report.get("passed", False):
        return END

    return "failed"
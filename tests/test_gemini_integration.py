
"""Integration tests for real Gemini workflow routing and tool calling."""

import os

import pytest
from dotenv import load_dotenv

from src.llm_agent import select_workflow_with_gemini
from src.workflow_engine import run_request

load_dotenv()


def require_gemini_key():
    """Skip real API tests when no Gemini key is configured."""
    if not (
        os.getenv("GOOGLE_API_KEY")
        or os.getenv("GEMINI_API_KEY")
    ):
        pytest.skip("Gemini API key is not configured.")


@pytest.mark.integration
def test_gemini_selects_inventory_workflow():
    """Verify Gemini selects WF001 for an inventory request."""
    require_gemini_key()

    decision = select_workflow_with_gemini(
        "Which products need restocking?"
    )

    assert decision is not None, (
        "Gemini returned no decision. Check API access and model settings."
    )
    assert decision.get("selected") is True, decision
    assert decision["workflow"]["Workflow_ID"] == "WF001"


@pytest.mark.integration
def test_langgraph_executes_gemini_selected_tool():
    """Verify the complete engine uses Gemini tool calling."""
    require_gemini_key()

    response = run_request("Which products need restocking?")

    assert response.get("status") == "success", response
    assert response.get("selected_workflow", {}).get("id") == "WF001"
    assert response.get("result", {}).get("count", -1) >= 0

    trace = response.get("execution_trace", [])
    assert any(
        item.get("routing_method") == "gemini_tool_call"
        for item in trace
    ), f"Gemini tool calling was not recorded in the trace: {trace}"

    assert any(
        item.get("step") == "Execute registered business tool"
        and item.get("status") == "success"
        for item in trace
    ), f"Business tool execution was not recorded: {trace}"

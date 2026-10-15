
from pathlib import Path

import pytest

from src.tool_registry import (
    TOOL_REGISTRY,
    _number,
    execute_tool,
    read_tabular_file,
)


# -------------------------------
# File validation tests
# -------------------------------

def test_missing_input_file_returns_clear_error(tmp_path):
    missing_file = tmp_path / "missing.csv"

    with pytest.raises(FileNotFoundError, match="Input file not found"):
        read_tabular_file(missing_file)


def test_unsupported_file_type_returns_clear_error(tmp_path):
    file_path = tmp_path / "input.txt"
    file_path.write_text("sample", encoding="utf-8")

    with pytest.raises(ValueError, match="Unsupported file type"):
        read_tabular_file(file_path)


def test_empty_file_returns_clear_error(tmp_path):
    file_path = tmp_path / "empty.csv"
    file_path.write_text("", encoding="utf-8")

    with pytest.raises(ValueError, match="Input file is empty"):
        read_tabular_file(file_path)


def test_invalid_numeric_value_is_not_silently_accepted():
    with pytest.raises(ValueError, match="Invalid numeric value"):
        _number("not-a-number", "current_stock")


# -------------------------------
# Workflow registration tests
# -------------------------------

def test_all_ten_workflows_are_registered():
    assert len(TOOL_REGISTRY) == 10

    expected_ids = {
        "WF001", "WF002", "WF003", "WF004", "WF005",
        "WF006", "WF007", "WF008", "WF009", "WF010",
    }

    assert set(TOOL_REGISTRY.keys()) == expected_ids


def test_every_registered_workflow_is_callable():
    for workflow_id, workflow_function in TOOL_REGISTRY.items():
        assert workflow_id.startswith("WF")
        assert callable(workflow_function)


def test_unknown_workflow_returns_clear_error():
    with pytest.raises(ValueError, match="Unknown workflow"):
        execute_tool("WF999", {})


def test_non_dictionary_context_is_rejected():
    with pytest.raises(ValueError, match="context must be a dictionary"):
        execute_tool("WF001", "invalid-context")


# -------------------------------
# Business workflow tests
# -------------------------------

def test_inventory_workflow_returns_dictionary():
    result = execute_tool("WF001", {})

    assert isinstance(result, dict)
    assert "status" in result


def test_price_validation_workflow_returns_dictionary():
    result = execute_tool("WF002", {})

    assert isinstance(result, dict)
    assert "status" in result


def test_order_status_requires_order_identifier():
    result = execute_tool("WF005", {})

    assert isinstance(result, dict)
    assert result.get("status") == "needs_input"


def test_campaign_workflow_requires_goal():
    result = execute_tool("WF007", {})

    assert isinstance(result, dict)
    assert result.get("status") == "needs_input"


def test_keyword_classification_returns_dictionary():
    result = execute_tool(
        "WF008",
        {"keywords": "buy shoes, best running shoes, shoe guide"},
    )

    assert isinstance(result, dict)
    assert "status" in result


def test_employee_assignment_returns_dictionary():
    result = execute_tool(
        "WF009",
        {"task": "Review product catalog"},
    )

    assert isinstance(result, dict)
    assert "status" in result


def test_performance_report_returns_dictionary():
    result = execute_tool("WF010", {})

    assert isinstance(result, dict)
    assert "status" in result

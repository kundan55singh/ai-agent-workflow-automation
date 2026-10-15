
"""LangGraph orchestration engine with Gemini tool calling."""

import re
import time
from typing import Any, TypedDict

from langgraph.graph import END, START, StateGraph

from src.gemini_tools import (
    ALLOWED_ARGUMENTS,
    execute_gemini_tool,
    request_gemini_tool_call,
)
from src.llm_agent import select_workflow_with_gemini
from src.tool_registry import execute_tool
from src.workflow_loader import load_excel_sheets
from src.workflow_router import select_workflow


class WorkflowState(TypedDict, total=False):
    user_request: str
    context: dict[str, Any]
    started_at: float
    workflows: list[dict[str, Any]]
    routing: dict[str, Any]
    workflow: dict[str, Any]
    business_context: dict[str, Any]
    result: dict[str, Any]
    trace: list[dict[str, Any]]
    output: dict[str, Any]


def _extract_context(request: str) -> dict[str, Any]:
    """Extract structured business inputs from natural-language requests."""
    context: dict[str, Any] = {"user_request": request}

    order_match = re.search(r"\bORD-[A-Za-z0-9-]+\b", request, re.I)
    if order_match:
        context["order_id"] = order_match.group(0).upper()

    email_match = re.search(
        r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
        request,
    )
    if email_match:
        context["customer_email"] = email_match.group(0)

    # Extract labelled campaign fields for deterministic fallback routing.
    field_patterns = {
        "campaign_goal": r"\b(?:campaign\s+)?(?:goal|objective)\s*:\s*([^.\n]+)",
        "target_audience": r"\btarget\s+audience\s*:\s*([^.\n]+)",
        "products": r"\bproducts?\s*:\s*([^.\n]+)",
        "promotion": r"\bpromotion\s*:\s*([^.\n]+)",
        "dates": r"\b(?:campaign\s+)?dates?\s*:\s*([^.\n]+)",
        "budget": r"\bbudget\s*:\s*([0-9]+(?:\.[0-9]+)?)",
    }

    for key, pattern in field_patterns.items():
        match = re.search(pattern, request, re.I)
        if match:
            value = match.group(1).strip().rstrip(",; ")
            if key == "products":
                context[key] = [
                    item.strip()
                    for item in re.split(r",|\band\b", value, flags=re.I)
                    if item.strip()
                ]
            else:
                context[key] = value

    return context

def load_workflows_node(state: WorkflowState) -> dict[str, Any]:
    """Load workflow definitions from Excel."""
    sheets = load_excel_sheets()
    workflows = sheets.get("Workflows", [])
    trace = list(state.get("trace", []))

    trace.append({
        "step": "Load workflow definitions",
        "status": "success",
        "detail": f"Loaded {len(workflows)} workflows from Excel.",
    })

    return {"workflows": workflows, "trace": trace}


def route_workflow_node(state: WorkflowState) -> dict[str, Any]:
    """Ask Gemini for a tool call, then use safe routing fallbacks."""
    request = state["user_request"]
    workflows = state.get("workflows", [])
    trace = list(state.get("trace", []))

    # First attempt: Gemini function/tool calling.
    try:
        decision = request_gemini_tool_call(request)
    except Exception as exc:
        print(
            f"Gemini tool calling unavailable: "
            f"{type(exc).__name__}: {str(exc)[:300]}"
        )
        decision = None

    if decision and decision.get("selected"):
        call = decision.get("tool_call", {})
        tool_name = call.get("name", "")
        workflow_id = (
            tool_name.removeprefix("run_").upper()
            if isinstance(tool_name, str)
            else ""
        )

        workflow = next(
            (
                item for item in workflows
                if str(item.get("Workflow_ID", "")) == workflow_id
            ),
            None,
        )

        if workflow is not None:
            routing = {
                "selected": True,
                "workflow": workflow,
                "score": None,
                "matched_terms": [],
                "routing_method": "gemini_tool_call",
                "tool_call": call,
            }

            trace.append({
                "step": "Select workflow",
                "status": "success",
                "detail": f"{workflow_id}: {workflow['Workflow_Name']}",
                "routing_method": "gemini_tool_call",
                "tool_name": tool_name,
            })

            return {
                "routing": routing,
                "workflow": workflow,
                "trace": trace,
            }

    # Second attempt: existing Gemini workflow selector.
    try:
        routing = select_workflow_with_gemini(request)
    except Exception as exc:
        print(
            f"Gemini workflow routing failed: "
            f"{type(exc).__name__}: {str(exc)[:300]}"
        )
        routing = None

    # Final fallback: deterministic local router.
    if routing is None:
        routing = select_workflow(request, workflows)
        routing["routing_method"] = "deterministic_fallback"

    if not routing.get("selected"):
        trace.append({
            "step": "Select workflow",
            "status": "needs_clarification",
            "detail": routing.get(
                "reason", "No suitable workflow was found."
            ),
        })

        return {
            "routing": routing,
            "trace": trace,
            "output": {
                "status": "needs_clarification",
                "request": request,
                "routing": routing,
                "execution_trace": trace,
                "result": None,
            },
        }

    workflow = routing["workflow"]

    trace.append({
        "step": "Select workflow",
        "status": "success",
        "detail": (
            f"{workflow['Workflow_ID']}: "
            f"{workflow['Workflow_Name']}"
        ),
        "matched_terms": routing.get("matched_terms", []),
        "routing_method": routing.get("routing_method", "unknown"),
    })

    return {"routing": routing, "workflow": workflow, "trace": trace}


def validate_inputs_node(state: WorkflowState) -> dict[str, Any]:
    """Prepare request and caller-provided context."""
    business_context = _extract_context(state["user_request"])
    business_context.update(state.get("context", {}))

    # Add arguments proposed by Gemini, without overwriting caller context.
    routing = state.get("routing", {})
    call = routing.get("tool_call", {})
    arguments = call.get("arguments", {}) if isinstance(call, dict) else {}

    if isinstance(arguments, dict):
        business_context.update(arguments)

    trace = list(state.get("trace", []))
    trace.append({
        "step": "Validate inputs",
        "status": "success",
        "detail": "Request context prepared for the selected workflow.",
    })

    return {"business_context": business_context, "trace": trace}


def execute_business_tool_node(
    state: WorkflowState,
) -> dict[str, Any]:
    """Execute the selected registered tool with controlled arguments."""
    workflow_id = state["workflow"]["Workflow_ID"]
    routing = state.get("routing", {})
    business_context = state.get("business_context", {})
    trace = list(state.get("trace", []))
    started = time.perf_counter()

    if routing.get("routing_method") == "gemini_tool_call":
        # Only pass arguments explicitly allowed for this workflow.
        allowed = ALLOWED_ARGUMENTS[workflow_id]
        tool_arguments = {
            key: value
            for key, value in business_context.items()
            if key in allowed
        }

        result = execute_gemini_tool(
            f"run_{workflow_id.lower()}",
            tool_arguments,
        )
        execution_method = "Gemini-selected registered tool"
    else:
        result = execute_tool(workflow_id, business_context)
        execution_method = "Registered Python tool via fallback routing"

    duration = round(time.perf_counter() - started, 4)

    trace.append({
        "step": "Execute registered business tool",
        "status": "success",
        "detail": (
            f"{execution_method} for {workflow_id} "
            f"completed in {duration}s."
        ),
    })

    return {"result": result, "trace": trace}


def finalize_node(state: WorkflowState) -> dict[str, Any]:
    """Build the structured result and execution trace."""
    workflow = state["workflow"]
    routing = state["routing"]
    result = state["result"]
    trace = list(state.get("trace", []))

    trace.append({
        "step": "Return result",
        "status": "success",
        "detail": "Structured workflow result generated.",
    })

    result_status = result.get("status")
    if result_status in {
        "needs_input", "not_found", "escalation_required"
    }:
        status = result_status
    elif result_status == "error":
        status = "error"
    else:
        status = "success"

    steps_text = str(workflow.get("Steps", ""))
    business_steps = [
        step.strip()
        for step in re.split(r"\s*(?:→|->|â†’)\s*", steps_text)
        if step.strip()
    ]

    output = {
        "status": status,
        "request": state["user_request"],
        "selected_workflow": {
            "id": workflow["Workflow_ID"],
            "name": workflow["Workflow_Name"],
            "expected_output": workflow["Expected_Output"],
        },
        "business_steps": business_steps,
        "decision_logic": workflow["Decision_Logic"],
        "routing": {
            "score": routing.get("score"),
            "matched_terms": routing.get("matched_terms", []),
            "routing_method": routing.get("routing_method", "unknown"),
        },
        "execution_trace": trace,
        "result": result,
        "execution_time_seconds": round(
            time.perf_counter() - state["started_at"], 4
        ),
    }

    return {"output": output, "trace": trace}


def _after_routing(state: WorkflowState) -> str:
    """Stop when routing needs clarification."""
    return "finish" if state.get("output") else "validate"


def build_workflow_graph():
    """Create the LangGraph orchestration pipeline."""
    graph = StateGraph(WorkflowState)

    graph.add_node("load_workflows", load_workflows_node)
    graph.add_node("route_workflow", route_workflow_node)
    graph.add_node("validate_inputs", validate_inputs_node)
    graph.add_node("execute_tool", execute_business_tool_node)
    graph.add_node("finalize", finalize_node)

    graph.add_edge(START, "load_workflows")
    graph.add_edge("load_workflows", "route_workflow")

    graph.add_conditional_edges(
        "route_workflow",
        _after_routing,
        {"validate": "validate_inputs", "finish": END},
    )

    graph.add_edge("validate_inputs", "execute_tool")
    graph.add_edge("execute_tool", "finalize")
    graph.add_edge("finalize", END)

    return graph.compile()


workflow_graph = build_workflow_graph()


def run_request(
    user_request: str,
    context: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Run a request through the compiled LangGraph."""
    started = time.perf_counter()

    initial_state: WorkflowState = {
        "user_request": user_request,
        "context": context or {},
        "started_at": started,
        "trace": [],
    }

    try:
        final_state = workflow_graph.invoke(initial_state)

        if final_state.get("output"):
            return final_state["output"]

        return {
            "status": "error",
            "request": user_request,
            "execution_trace": final_state.get("trace", []),
            "error": "The graph finished without producing an output.",
        }

    except Exception as exc:
        trace = list(initial_state.get("trace", []))
        trace.append({
            "step": "Handle error",
            "status": "error",
            "detail": f"{type(exc).__name__}: {exc}",
        })

        return {
            "status": "error",
            "request": user_request,
            "execution_trace": trace,
            "error": str(exc),
            "execution_time_seconds": round(
                time.perf_counter() - started, 4
            ),
        }

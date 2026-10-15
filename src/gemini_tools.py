
"""Gemini-compatible definitions for registered business workflow tools."""

import json
import os
from typing import Any

from dotenv import load_dotenv
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_google_genai import ChatGoogleGenerativeAI

from src.tool_registry import TOOL_REGISTRY

load_dotenv()

# Gemini can request only these registered business tools.
TOOL_DESCRIPTIONS = {
    "WF001": "Check inventory and identify products below their restock threshold.",
    "WF002": "Compare internal and vendor prices and flag differences above a threshold.",
    "WF003": "Validate a vendor spreadsheet and identify invalid rows.",
    "WF004": "Generate a product description and SEO metadata from supplied product details.",
    "WF005": "Look up the status of a customer order using its order ID or email.",
    "WF006": "Detect likely duplicate products in a product catalog.",
    "WF007": "Prepare a marketing campaign brief from supplied campaign details.",
    "WF008": "Classify SEO keywords by search intent and recommend target pages.",
    "WF009": "Recommend an available employee based on skills and workload.",
    "WF010": "Summarize workflow execution performance and identify failures.",
}

# Limit the arguments the model can request.
ALLOWED_ARGUMENTS = {
    "WF001": {"inventory_file"},
    "WF002": {"prices_file", "threshold_percent"},
    "WF003": {"vendor_file", "write_files"},
    "WF004": {
        "product_name", "product", "category", "material",
        "color", "target_audience",
    },
    "WF005": {"order_id", "customer_email"},
    "WF006": {"products_file"},
    "WF007": {
        "campaign_goal", "dates", "products",
        "target_audience", "promotion",
    },
    "WF008": {"keywords", "keyword_file", "category"},
    "WF009": {
        "task_description", "required_skills", "priority", "deadline",
    },
    "WF010": {"slow_threshold_seconds"},
}


def get_gemini_tool_schemas() -> list[dict[str, Any]]:
    """Create JSON schemas describing the registered workflow tools."""
    schemas = []

    for workflow_id, function in TOOL_REGISTRY.items():
        schemas.append({
            "name": f"run_{workflow_id.lower()}",
            "description": TOOL_DESCRIPTIONS[workflow_id],
            "parameters": {
                "type": "object",
                "properties": {
                    argument: {"type": "string"}
                    for argument in sorted(ALLOWED_ARGUMENTS[workflow_id])
                },
                "required": [],
            },
        })

    return schemas


def execute_gemini_tool(
    tool_name: str,
    arguments: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate a Gemini tool request and execute a registered Python tool."""
    arguments = arguments or {}

    if not isinstance(arguments, dict):
        raise ValueError("Tool arguments must be a JSON object.")

    if not tool_name.startswith("run_wf"):
        raise ValueError("Unknown tool name.")

    workflow_id = tool_name.removeprefix("run_").upper()

    if workflow_id not in TOOL_REGISTRY:
        raise ValueError(f"Unregistered workflow tool: {tool_name}")

    allowed = ALLOWED_ARGUMENTS[workflow_id]
    unexpected = set(arguments) - allowed

    if unexpected:
        raise ValueError(
            f"Unsupported arguments for {workflow_id}: "
            f"{sorted(unexpected)}"
        )

    # File paths are supplied by the application, not trusted to the model.
    file_arguments = {
        "inventory_file", "prices_file", "vendor_file",
        "products_file", "keyword_file",
    }
    for key in file_arguments & set(arguments):
        if not isinstance(arguments[key], str):
            raise ValueError(f"{key} must be a string path.")

    context = dict(arguments)

    # Convert common numeric and boolean arguments safely.
    if "threshold_percent" in context:
        context["threshold_percent"] = float(context["threshold_percent"])
    if "slow_threshold_seconds" in context:
        context["slow_threshold_seconds"] = float(
            context["slow_threshold_seconds"]
        )
    if "write_files" in context:
        value = context["write_files"]
        if isinstance(value, str):
            if value.casefold() not in {"true", "false"}:
                raise ValueError("write_files must be true or false.")
            context["write_files"] = value.casefold() == "true"

    return TOOL_REGISTRY[workflow_id](context)


def request_gemini_tool_call(
    user_request: str,
) -> dict[str, Any]:
    """Ask Gemini which registered business tool best fits a request."""
    api_key = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Gemini API key is not configured.")

    model_name = os.getenv("GEMINI_MODEL", "gemini-2.5-flash").strip()

    model = ChatGoogleGenerativeAI(
        model=model_name,
        google_api_key=api_key,
        max_retries=1,
        timeout=30,
    ).bind_tools(get_gemini_tool_schemas())

    response = model.invoke([
        SystemMessage(content=(
            "Choose a registered business tool only when it matches the user's "
            "request. Never invent data or execute code. If required business "
            "inputs are missing, ask for them instead of inventing values. "
            "The application will validate and execute any requested tool."
        )),
        HumanMessage(content=user_request),
    ])

    tool_calls = getattr(response, "tool_calls", []) or []

    if not tool_calls:
        return {
            "selected": False,
            "message": str(response.content),
            "tool_call": None,
        }

    call = tool_calls[0]
    name = call.get("name", "")
    arguments = call.get("args", {})

    # Validate the tool request before returning it to the application.
    if name not in {schema["name"] for schema in get_gemini_tool_schemas()}:
        raise ValueError(f"Gemini requested an unknown tool: {name}")

    if not isinstance(arguments, dict):
        raise ValueError("Gemini returned invalid tool arguments.")

    return {
        "selected": True,
        "tool_call": {
            "name": name,
            "arguments": arguments,
        },
        "message": "",
    }

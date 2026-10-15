
import re
from difflib import SequenceMatcher
from typing import Any

from src.workflow_loader import load_excel_sheets


def get_workflows() -> list[dict[str, Any]]:
    sheets = load_excel_sheets()
    workflows = sheets.get("Workflows", [])

    if not workflows:
        raise ValueError("No workflows found in the Workflows sheet.")

    required = {
        "Workflow_ID",
        "Workflow_Name",
        "Trigger",
        "Inputs",
        "Steps",
        "Decision_Logic",
        "Tools_Required",
        "Expected_Output",
    }

    for row in workflows:
        missing = required - row.keys()
        if missing:
            raise ValueError(
                f"Excel row {row.get('_excel_row')} is missing: {sorted(missing)}"
            )

    return workflows


def _tokens(text: str) -> set[str]:
    words = set(re.findall(r"[a-z0-9]+", text.lower()))

    # Simple normalization for common English plurals and verb forms.
    normalized = set()
    for word in words:
        if word.endswith("ies") and len(word) > 4:
            word = word[:-3] + "y"
        elif word.endswith("ing") and len(word) > 5:
            word = word[:-3]
        elif word.endswith("s") and len(word) > 3:
            word = word[:-1]
        normalized.add(word)

    return normalized


def _similarity(left: str, right: str) -> float:
    a = " ".join(sorted(_tokens(left)))
    b = " ".join(sorted(_tokens(right)))
    if not a or not b:
        return 0.0

    sequence_score = SequenceMatcher(None, a, b).ratio()
    left_tokens, right_tokens = _tokens(left), _tokens(right)
    overlap = len(left_tokens & right_tokens) / max(len(left_tokens), 1)

    return max(sequence_score, overlap)


def select_workflow(
    user_request: str,
    workflows: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    if not user_request or not user_request.strip():
        return {
            "selected": False,
            "reason": "Please describe the task you want to perform.",
        }

    sheets = load_excel_sheets()
    workflows = workflows or sheets.get("Workflows", [])
    examples = sheets.get("Test_Questions", [])

    if not workflows:
        return {"selected": False, "reason": "No workflows are configured."}

    workflow_by_id = {str(row["Workflow_ID"]): row for row in workflows}
    request = user_request.strip()

    # Prefer the provided test questions when the request closely matches one.
    example_scores = []
    for example in examples:
        score = _similarity(request, str(example.get("Test_Request", "")))
        example_scores.append((score, str(example.get("Workflow_ID", ""))))

    if example_scores:
        example_scores.sort(reverse=True)
        best_score, best_id = example_scores[0]
        second_score = example_scores[1][0] if len(example_scores) > 1 else 0.0

        if best_score >= 0.78 and (
            best_score - second_score >= 0.04 or best_score >= 0.95
        ):
            if best_id in workflow_by_id:
                return {
                    "selected": True,
                    "workflow": workflow_by_id[best_id],
                    "score": round(best_score, 3),
                    "matched_terms": ["Excel test-question match"],
                    "routing_method": "test_question_similarity",
                }

    # Fallback: compare against the workflow's name, trigger and description.
    scored = []

    for workflow in workflows:
        candidates = [
            str(workflow.get("Workflow_Name", "")),
            str(workflow.get("Trigger", "")),
            str(workflow.get("Steps", "")),
            str(workflow.get("Inputs", "")),
        ]

        score = max(_similarity(request, candidate) for candidate in candidates)
        scored.append((score, workflow))

    scored.sort(key=lambda item: item[0], reverse=True)

    best_score, best_workflow = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0

    if best_score < 0.30:
        return {
            "selected": False,
            "reason": "No sufficiently matching workflow was found.",
            "available_workflows": [
                row["Workflow_Name"] for row in workflows
            ],
        }

    if len(scored) > 1 and best_score - second_score < 0.04:
        return {
            "selected": False,
            "reason": "The request is ambiguous; please clarify the intended task.",
            "candidates": [
                {
                    "id": row["Workflow_ID"],
                    "name": row["Workflow_Name"],
                    "score": round(score, 3),
                }
                for score, row in scored
                if best_score - score < 0.04
            ],
        }

    return {
        "selected": True,
        "workflow": best_workflow,
        "score": round(best_score, 3),
        "matched_terms": [],
        "routing_method": "workflow_description_similarity",
    }


"""Interactive CLI for the AI Agent Workflow Automation project."""

import json
import sys

from dotenv import load_dotenv

from src.workflow_engine import run_request


def display_result(result: dict) -> None:
    """Display workflow selection, execution steps and final result."""

    print("\n" + "=" * 65)
    print("AI AGENT WORKFLOW AUTOMATION")
    print("=" * 65)

    print(f"\nStatus: {result.get('status', 'unknown')}")

    selected = result.get("selected_workflow")
    if selected:
        print("\nSelected Workflow")
        print("-" * 65)
        print(f"ID:   {selected.get('id', 'N/A')}")
        print(f"Name: {selected.get('name', 'N/A')}")
        print(f"Reason: {result.get('routing', {}).get('routing_method', 'N/A')}")

    print("\nExecution Steps")
    print("-" * 65)

    for index, step in enumerate(result.get("execution_trace", []), start=1):
        print(
            f"{index}. [{step.get('status', 'unknown')}] "
            f"{step.get('step', 'Step')}"
        )
        detail = step.get("detail")
        if detail:
            print(f"   {detail}")

    print("\nFinal Result")
    print("-" * 65)

    final_result = result.get("result")
    if final_result is not None:
        print(json.dumps(final_result, indent=2, ensure_ascii=False, default=str))
    else:
        print(result.get("error") or result.get("routing", {}).get(
            "reason", "No result was returned."
        ))

    if result.get("execution_time_seconds") is not None:
        print(
            f"\nExecution time: "
            f"{result['execution_time_seconds']} seconds"
        )

    print("=" * 65)


def main() -> None:
    load_dotenv()

    print("=" * 65)
    print("Welcome to AI Agent Workflow Automation")
    print("Powered by Gemini, LangGraph, Python and Excel")
    print("=" * 65)
    print("Enter a business request in natural language.")
    print("Type 'exit' or 'quit' to close the application.\n")

    while True:
        try:
            user_request = input("You: ").strip()

            if user_request.lower() in {"exit", "quit"}:
                print("Application closed. Goodbye!")
                break

            if not user_request:
                print("Please enter a request.\n")
                continue

            result = run_request(user_request)
            display_result(result)

        except KeyboardInterrupt:
            print("\nApplication interrupted. Goodbye!")
            break

        except Exception as exc:
            print(f"\nApplication error: {type(exc).__name__}: {exc}")
            print("Please try again with a valid business request.\n")


if __name__ == "__main__":
    main()

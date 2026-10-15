# AI Agent Workflow Automation

A reusable Python-based AI agent system that reads business workflow definitions from Excel, identifies the appropriate workflow from a natural-language request, and executes the corresponding registered business tool.

## Features

- Excel-driven workflow definitions.
- Gemini-powered workflow and tool selection.
- LangGraph-based workflow execution pipeline.
- Deterministic routing fallback.
- Centralized registry for 10 business workflows.
- Structured results, execution traces, and error handling.
- Input validation and automated tests.

## Architecture

```text
User Request
     |
     v
Excel Workflow Loader
     |
     v
Gemini Tool Selection
     |
     v
Workflow Router
     |
     v
Input Validation
     |
     v
Registered Business Tool
     |
     v
Structured Result and Execution Trace
```

The workflow engine coordinates these stages. Gemini assists with selecting a suitable registered tool, while the tool registry controls which business functions can be executed.

## Supported Workflows

| Workflow ID | Workflow |
|---|---|
| WF001 | Inventory Restock Check |
| WF002 | Price Validation |
| WF003 | Vendor File Processing |
| WF004 | Product Content Generation |
| WF005 | Order Status Lookup |
| WF006 | Duplicate Detection |
| WF007 | Marketing Campaign Brief |
| WF008 | Keyword Classification |
| WF009 | Employee Task Assignment |
| WF010 | Workflow Performance Monitoring |

Workflow definitions are maintained in `data/business_workflows.xlsx`.

## Technology Stack

- Python
- Google Gemini
- LangChain and LangGraph
- Excel processing with openpyxl
- pytest
- python-dotenv

## Project Structure

```text
AI Workflow/
├── app.py
├── data/
│   └── business_workflows.xlsx
├── examples/
├── outputs/
├── src/
│   ├── gemini_tools.py
│   ├── llm_agent.py
│   ├── tool_registry.py
│   ├── workflow_engine.py
│   ├── workflow_loader.py
│   └── workflow_router.py
├── tests/
│   ├── test_workflows.py
│   └── test_gemini_integration.py
├── .env.example
├── .gitignore
├── pytest.ini
└── requirements.txt
```

## Setup

### 1. Create and activate a virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2. Install dependencies

```powershell
python -m pip install -r requirements.txt
```

### 3. Configure environment variables

Create a local `.env` file from the example:

```powershell
Copy-Item .env.example .env
```

Add your valid Gemini API key to `.env` and configure the model supported by your Gemini API account.

```dotenv
GEMINI_API_KEY=your_gemini_api_key_here
GEMINI_MODEL=your_supported_gemini_model
```

Never commit `.env` or expose your API key publicly.

### 4. Run the application

```powershell
python app.py
```

Enter a business request when prompted. Type `exit` or `quit` to close the application.

## Example Requests

```text
Which products need restocking?
Check which products have incorrect prices.
Process the vendor file.
Generate product content for a new product.
Check the status of order ORD-1001.
Find duplicate products.
Create a marketing campaign brief.
Classify these keywords.
Assign an employee to this task.
Check workflow performance.
```

Requests requiring additional information or input files may need the appropriate context or file path.

## Run Tests

Run the complete test suite:

```powershell
python -m pytest tests -v
```

The current development environment has passed all 17 tests, including two Gemini integration tests. Integration tests require valid API configuration and network access.

## Extending the System

To add another workflow:

1. Add its definition to the Excel workbook.
2. Implement its business logic in `src/tool_registry.py`.
3. Register the new workflow ID in `TOOL_REGISTRY`.
4. Define its allowed tool arguments and connect the tool schema/routing as required.
5. Add unit tests and, where applicable, integration tests.

This design separates workflow definitions, routing, orchestration, and business logic to make the system easier to extend and maintain.

## Security Notes

- Keep API keys in local environment variables.
- Do not commit `.env`, virtual environments, or private credentials.
- Validate external file inputs before processing.
- Review file paths and tool arguments before executing workflows on production data.

## Project Status

The application successfully executed the inventory restock workflow in the development environment. The full automated test suite passed with 17 tests.

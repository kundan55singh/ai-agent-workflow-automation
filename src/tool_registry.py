
import csv
import math
import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


MAX_INPUT_FILE_BYTES = 10 * 1024 * 1024

INVENTORY = [
    {"sku": "SKU-101", "name": "Wireless Mouse", "current_stock": 4, "reorder_threshold": 10},
    {"sku": "SKU-102", "name": "Keyboard", "current_stock": 25, "reorder_threshold": 10},
    {"sku": "SKU-103", "name": "USB-C Cable", "current_stock": 3, "reorder_threshold": 15},
]

PRICES = [
    {"sku": "SKU-101", "product_name": "Wireless Mouse", "internal_price": 499, "vendor_price": 520},
    {"sku": "SKU-102", "product_name": "Keyboard", "internal_price": 899, "vendor_price": 899},
    {"sku": "SKU-103", "product_name": "USB-C Cable", "internal_price": 199, "vendor_price": 250},
]

ORDERS = [
    {
        "order_id": "ORD-1001",
        "email": "customer@example.com",
        "status": "Shipped",
        "shipment_status": "In transit",
        "tracking_id": "TRK-1001",
        "items": ["Wireless Keyboard", "Wireless Mouse"],
    },
    {
        "order_id": "ORD-1002",
        "email": "buyer@example.com",
        "status": "Processing",
        "shipment_status": "Not shipped",
        "tracking_id": "",
        "items": ["USB-C Cable"],
    },
]

PRODUCTS = [
    {"sku": "SKU-101", "name": "Wireless Mouse", "category": "Electronics"},
    {"sku": "SKU-102", "name": "Mechanical Keyboard", "category": "Electronics"},
]

EMPLOYEES = [
    {
        "name": "Aarav",
        "skills": "inventory, reporting, python, excel",
        "active_tasks": 2,
        "available": True,
        "max_tasks": 5,
    },
    {
        "name": "Diya",
        "skills": "marketing, content, seo",
        "active_tasks": 1,
        "available": True,
        "max_tasks": 4,
    },
    {
        "name": "Kabir",
        "skills": "data, reporting, analytics, python, excel",
        "active_tasks": 3,
        "available": True,
        "max_tasks": 5,
    },
]

EXECUTION_HISTORY = [
    {"workflow_id": "WF001", "duration_seconds": 1.2, "status": "success"},
    {"workflow_id": "WF002", "duration_seconds": 2.4, "status": "success"},
    {"workflow_id": "WF003", "duration_seconds": 6.2, "status": "failed"},
]


# ==================================================
# FILE READING AND VALIDATION
# ==================================================

def read_tabular_file(file_path: str | Path) -> list[dict[str, Any]]:
    """Read CSV or Excel files safely and return rows as dictionaries."""
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    if not path.is_file():
        raise ValueError(f"Input path is not a file: {path}")

    if path.suffix.lower() not in {".csv", ".xlsx", ".xlsm"}:
        raise ValueError(f"Unsupported file type: {path.suffix}")

    if path.stat().st_size == 0:
        raise ValueError(f"Input file is empty: {path}")

    if path.stat().st_size > MAX_INPUT_FILE_BYTES:
        raise ValueError("Input file exceeds the 10 MB size limit")

    try:
        if path.suffix.lower() == ".csv":
            with path.open("r", encoding="utf-8-sig", newline="") as file:
                reader = csv.DictReader(file)

                if not reader.fieldnames:
                    raise ValueError("Input file has no header row")

                headers = [
                    str(header).strip() if header is not None else ""
                    for header in reader.fieldnames
                ]

                if any(not header for header in headers):
                    raise ValueError("Input file contains an empty column header")

                if len(headers) != len(set(headers)):
                    raise ValueError("Input file contains duplicate column headers")

                rows = []

                for row in reader:
                    if not row or all(
                        value is None or not str(value).strip()
                        for value in row.values()
                    ):
                        continue

                    rows.append({
                        str(key).strip(): (
                            value.strip() if isinstance(value, str) else value
                        )
                        for key, value in row.items()
                        if key is not None
                    })

        else:
            workbook = load_workbook(
                path,
                read_only=True,
                data_only=True,
            )

            try:
                worksheet = workbook.active
                iterator = worksheet.iter_rows(values_only=True)
                header_row = next(iterator, None)

                if not header_row:
                    raise ValueError("Input file has no header row")

                headers = [
                    str(header).strip() if header is not None else ""
                    for header in header_row
                ]

                if any(not header for header in headers):
                    raise ValueError("Input file contains an empty column header")

                if len(headers) != len(set(headers)):
                    raise ValueError("Input file contains duplicate column headers")

                rows = []

                for values in iterator:
                    if not values or all(
                        value is None or not str(value).strip()
                        for value in values
                    ):
                        continue

                    rows.append(dict(zip(headers, values)))

            finally:
                workbook.close()

    except UnicodeDecodeError as exc:
        raise ValueError("CSV file must use UTF-8 encoding") from exc

    except InvalidFileException as exc:
        raise ValueError("Invalid or corrupted Excel file") from exc

    except (csv.Error, OSError) as exc:
        raise ValueError(f"Unable to read input file: {exc}") from exc

    if not rows:
        raise ValueError("Input file contains no data rows")

    return rows


# ==================================================
# COMMON HELPERS
# ==================================================

def _number(value: Any, field_name: str) -> float:
    """Convert input to a finite number with clear error handling."""
    try:
        if value is None or isinstance(value, bool) or str(value).strip() == "":
            raise ValueError

        number = float(value)

        if not math.isfinite(number):
            raise ValueError

        return number

    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError(
            f"Invalid numeric value for '{field_name}': {value!r}"
        ) from exc


def _rows(
    context: dict,
    file_key: str,
    fallback: list[dict],
) -> list[dict]:
    """Read records from a file or use the workflow's demo data."""
    file_path = context.get(file_key)

    if file_path:
        return read_tabular_file(file_path)

    records = context.get(file_key.removesuffix("_file"), fallback)

    if not isinstance(records, list):
        raise ValueError(f"{file_key} data must be a list")

    if any(not isinstance(record, dict) for record in records):
        raise ValueError(f"Every record in {file_key} data must be a dictionary")

    return records


def _field(row: dict, *names: str, default=None):
    """Find a field using case-insensitive column names."""
    normalized = {
        str(key).strip().lower().replace(" ", "_"): value
        for key, value in row.items()
    }

    for name in names:
        key = name.strip().lower().replace(" ", "_")

        if key in normalized:
            return normalized[key]

    return default


def _success(data: Any, message: str = "Workflow completed successfully") -> dict:
    return {
        "status": "success",
        "message": message,
        "data": data,
    }


def _needs_input(message: str, required_fields: list[str]) -> dict:
    return {
        "status": "needs_input",
        "message": message,
        "required_fields": required_fields,
    }


# ==================================================
# WF001: INVENTORY RESTOCK CHECK
# ==================================================

def inventory_restock(context: dict) -> dict:
    records = _rows(context, "inventory_file", INVENTORY)
    restock = []

    for row in records:
        stock = _number(
            _field(row, "current_stock", "stock", "quantity"),
            "current_stock",
        )
        threshold = _number(
            _field(row, "reorder_threshold", "minimum_stock", "threshold"),
            "reorder_threshold",
        )

        if stock < 0 or threshold < 0:
            raise ValueError("Stock and reorder threshold cannot be negative")

        if stock < threshold:
            restock.append({
                "sku": _field(row, "sku", "product_id", default="Unknown"),
                "product": _field(row, "name", "product_name", default="Unknown"),
                "current_stock": stock,
                "reorder_threshold": threshold,
                "recommended_order_quantity": max(
                    0, int(threshold * 2 - stock)
                ),
            })

    return {
        "status": "success",
        "message": f"Found {len(restock)} product(s) requiring a restock check",
        "count": len(restock),
        "data": restock,
    }


# ==================================================
# WF002: PRODUCT PRICE VALIDATION
# ==================================================

def validate_prices(context: dict) -> dict:
    records = _rows(context, "prices_file", PRICES)

    threshold = _number(
        context.get("threshold_percent", 10),
        "threshold_percent",
    )

    if threshold < 0:
        raise ValueError("threshold_percent cannot be negative")

    discrepancies = []

    for row in records:
        internal = _number(
            _field(row, "internal_price", "our_price"),
            "internal_price",
        )
        vendor = _number(
            _field(row, "vendor_price", "supplier_price"),
            "vendor_price",
        )

        if internal < 0 or vendor < 0:
            raise ValueError("Product prices cannot be negative")

        difference = abs(vendor - internal)
        percent = difference / internal * 100 if internal else (
            100.0 if vendor else 0.0
        )

        if percent > threshold:
            discrepancies.append({
                "sku": _field(row, "sku", default="Unknown"),
                "product": _field(row, "product_name", "product", "name", default="Unknown"),
                "internal_price": internal,
                "vendor_price": vendor,
                "difference_percent": round(percent, 2),
            })

    return _success(discrepancies, "Price validation completed")


# ==================================================
# WF003: VENDOR FILE PROCESSING
# ==================================================

def process_vendor_file(context: dict) -> dict:
    file_path = context.get("vendor_file")

    if not file_path:
        return _needs_input(
            "Provide a vendor CSV or Excel file to process.",
            ["vendor_file"],
        )

    records = read_tabular_file(file_path)
    output_dir = Path(context.get("output_dir", "outputs"))
    output_dir.mkdir(parents=True, exist_ok=True)

    valid_records = []
    invalid_records = []

    for index, row in enumerate(records, start=2):
        normalized = {
            str(key).strip().lower().replace(" ", "_"): value
            for key, value in row.items()
        }

        sku = normalized.get("sku")
        product_name = normalized.get("product_name", normalized.get("name"))

        if sku is None or not str(sku).strip():
            invalid_records.append({
                "row": index,
                "error": "SKU is required",
                "data": str(row),
            })
            continue

        if product_name is None or not str(product_name).strip():
            invalid_records.append({
                "row": index,
                "error": "Product name is required",
                "data": str(row),
            })
            continue

        cleaned_row = {
            str(key).strip().lower().replace(" ", "_"): value
            for key, value in row.items()
        }
        cleaned_row["sku"] = str(sku).strip()
        cleaned_row["product_name"] = str(product_name).strip()
        valid_records.append(cleaned_row)

    cleaned_path = output_dir / "cleaned_vendor.csv"
    invalid_path = output_dir / "invalid_vendor_rows.csv"

    fieldnames = list(valid_records[0].keys()) if valid_records else [
        "sku", "product_name"
    ]
    with cleaned_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(valid_records)

    with invalid_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["row", "error", "data"], extrasaction="ignore"
        )
        writer.writeheader()
        writer.writerows(invalid_records)

    return _success({
        "total_rows": len(records),
        "valid_rows": len(valid_records),
        "invalid_rows": len(invalid_records),
        "cleaned_file": str(cleaned_path),
        "invalid_rows_file": str(invalid_path),
        "invalid_row_details": invalid_records,
    }, "Vendor file processing completed")

def generate_product_content(context: dict) -> dict:
    name = str(context.get("product_name", "")).strip()

    if not name:
        return _needs_input(
            "Provide the product name.",
            ["product_name"],
        )

    audience = str(
        context.get("target_audience", "everyday customers")
    ).strip()

    material = str(context.get("material", "")).strip()
    color = str(context.get("color", "")).strip()
    category = str(context.get("category", "")).strip()
    features = context.get("features", [])

    if isinstance(features, str):
        features = [
            item.strip()
            for item in features.split(",")
            if item.strip()
        ]

    if not isinstance(features, list):
        raise ValueError(
            "features must be a list or comma-separated string"
        )

    details = []
    if category:
        details.append(f"category: {category}")
    if material:
        details.append(f"made with {material}")
    if color:
        details.append(f"available in {color}")
    details.extend(str(item).strip() for item in features if str(item).strip())

    detail_text = "; ".join(details) or "practical design and everyday usability"

    short_description = (
        f"{name} for {audience}, featuring {detail_text}."
    )

    description = (
        f"Discover the {name}. Designed for {audience}, this product "
        f"offers {detail_text}. It is a practical option for customers "
        f"looking for quality and convenience in everyday use."
    )

    seo_title = f"{name} | Features and Benefits"
    seo_description = (
        f"Explore {name} for {audience}. Discover its {detail_text}."
    )

    return _success({
        "product_name": name,
        "category": category or "Not specified",
        "material": material or "Not specified",
        "color": color or "Not specified",
        "target_audience": audience,
        "short_description": short_description,
        "description": description,
        "seo_title": seo_title,
        "seo_description": seo_description,
    }, "Product content generated")


# ==================================================
# WF005: CUSTOMER ORDER STATUS
# ==================================================

def get_order_status(context: dict) -> dict:
    order_id = str(context.get("order_id", "")).strip()
    email = str(context.get("email") or context.get("customer_email") or "").strip()

    if not order_id and not email:
        return _needs_input(
            "Provide an order ID or customer email.",
            ["order_id", "email"],
        )

    if email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
        raise ValueError("Invalid email address")

    matches = [
        order for order in ORDERS
        if (
            order_id
            and order["order_id"].lower() == order_id.lower()
        ) or (
            email
            and order["email"].lower() == email.lower()
        )
    ]

    if not matches:
        return {
            "status": "not_found",
            "message": "No matching order was found in the demo dataset",
        }

    return _success(matches, "Order lookup completed")


# ==================================================
# WF006: DUPLICATE PRODUCT DETECTION
# ==================================================

def detect_duplicates(context: dict) -> dict:
    from difflib import SequenceMatcher

    records = _rows(context, "products_file", PRODUCTS)
    seen_skus = {}
    seen_names = {}
    duplicates = []
    similarity_threshold = 0.85

    for row in records:
        sku_original = str(
            _field(row, "sku", "product_id", default="")
        ).strip()
        name_original = str(
            _field(row, "name", "product_name", default="")
        ).strip()

        sku = sku_original.lower()
        name = name_original.lower()
        normalized_name = " ".join(name.split())

        duplicate_of = None
        match_type = None
        confidence = 0.0

        if sku and sku in seen_skus:
            duplicate_of = seen_skus[sku]
            match_type = "exact_sku"
            confidence = 1.0

        elif normalized_name and normalized_name in seen_names:
            duplicate_of = seen_names[normalized_name]
            match_type = "exact_name"
            confidence = 1.0

        elif normalized_name:
            for existing_name, existing_sku in seen_names.items():
                score = SequenceMatcher(
                    None, normalized_name, existing_name
                ).ratio()

                if score >= similarity_threshold and score > confidence:
                    confidence = score
                    duplicate_of = existing_sku
                    match_type = "similar_name"

        if duplicate_of:
            duplicates.append({
                "sku": sku_original,
                "name": name_original,
                "duplicate_of": duplicate_of,
                "match_type": match_type,
                "confidence": round(confidence, 3),
            })
        else:
            if sku:
                seen_skus[sku] = sku_original
            if normalized_name:
                seen_names[normalized_name] = sku_original or name_original

    return _success(
        {
            "total_products_checked": len(records),
            "duplicate_count": len(duplicates),
            "similarity_threshold": similarity_threshold,
            "duplicates": duplicates,
        },
        f"Checked {len(records)} products; found {len(duplicates)} duplicate(s)",
    )


# ==================================================
# WF007: MARKETING CAMPAIGN BRIEF
# ==================================================

def create_campaign_brief(context: dict) -> dict:
    goal = str(context.get("campaign_goal", "")).strip()

    if not goal:
        return _needs_input(
            "Provide a campaign goal.",
            ["campaign_goal"],
        )

    audience = str(
        context.get("target_audience", "general audience")
    ).strip()

    channel = str(
        context.get("channel", "social media")
    ).strip()

    products = context.get("products", [])
    if isinstance(products, str):
        products = [
            item.strip()
            for item in re.split(r",|\\band\\b", products, flags=re.I)
            if item.strip()
        ]
    if not isinstance(products, list):
        raise ValueError("products must be a list or comma-separated string")

    promotion = str(context.get("promotion", "")).strip()
    dates = str(context.get("dates", "")).strip()

    key_message = (
        f"{promotion} on {', '.join(str(p) for p in products)}"
        if promotion and products
        else promotion or (
            f"Discover {', '.join(str(p) for p in products)}"
            if products else f"Communicate a clear benefit to {audience}"
        )
    )

    brief = {
        "objective": goal,
        "target_audience": audience,
        "products": products,
        "promotion": promotion or "No promotion specified",
        "channel": channel,
        "campaign_dates": dates or "Not specified",
        "key_message": key_message,
        "call_to_action": "Shop now" if promotion else "Learn more",
        "timeline": [
            {"phase": "Preparation", "tasks": [
                "Finalize campaign creative and product details",
                "Prepare landing pages and tracking"
            ]},
            {"phase": "Launch", "tasks": [
                "Publish campaign content",
                "Activate promotion and monitor engagement"
            ]},
            {"phase": "Optimization", "tasks": [
                "Review clicks, conversions and sales",
                "Adjust messaging and channel performance"
            ]},
            {"phase": "Wrap-up", "tasks": [
                "Summarize campaign results",
                "Document learnings and next steps"
            ]}
        ],
        "execution_checklist": [
            "Confirm product availability and promotion terms",
            "Approve campaign copy and creatives",
            "Verify landing pages and tracking links",
            "Schedule campaign content",
            "Monitor performance and budget",
            "Prepare final performance report"
        ]
    }

    budget = context.get("budget")
    if budget not in (None, ""):
        amount = _number(budget, "budget")
        if amount < 0:
            raise ValueError("budget cannot be negative")
        brief["budget"] = amount

    return _success(brief, "Marketing campaign brief created")


# ==================================================
# WF008: SEO KEYWORD CLASSIFICATION
# ==================================================

def classify_keywords(context: dict) -> dict:
    keywords = context.get("keywords", [])

    if isinstance(keywords, str):
        keywords = [
            item.strip()
            for item in keywords.split(",")
            if item.strip()
        ]

    if not keywords and context.get("keyword_file"):
        records = read_tabular_file(context["keyword_file"])
        keywords = [
            str(_field(row, "keyword", "query", default="")).strip()
            for row in records
        ]
        keywords = [keyword for keyword in keywords if keyword]

    if not isinstance(keywords, list):
        raise ValueError(
            "keywords must be a list or comma-separated string"
        )

    if not keywords:
        return _needs_input(
            "Provide keywords to classify.",
            ["keywords"],
        )

    classified = []
    seen = set()

    for keyword in keywords:
        text = str(keyword).strip()
        lower = " ".join(text.lower().split())

        if not text or lower in seen:
            continue
        seen.add(lower)

        if any(term in lower for term in (
            "buy ", "price", "discount", "order ", "cheap",
            "purchase", "shop "
        )):
            intent = "Transactional"
        elif any(term in lower for term in (
            "best ", "versus", " vs ", "review",
            "compare", "top "
        )):
            intent = "Commercial"
        elif any(term in lower for term in (
            "how ", "what ", "why ", "guide",
            "meaning", "tutorial", "how to"
        )):
            intent = "Informational"
        else:
            intent = "Commercial"

        if any(term in lower for term in (
            "keyboard", "mouse", "laptop", "monitor"
        )):
            category = "Electronics"
        elif any(term in lower for term in (
            "cable", "hub", "adapter", "charger"
        )):
            category = "Accessories"
        else:
            category = "General"

        if intent == "Transactional":
            priority = "High"
            target_page = "Product page"
        elif intent == "Commercial":
            priority = "High" if any(
                term in lower for term in ("best ", "compare", "review")
            ) else "Medium"
            target_page = "Comparison or buying guide"
        else:
            priority = "Medium"
            target_page = "Blog or educational guide"

        classified.append({
            "keyword": text,
            "intent": intent,
            "category": category,
            "priority": priority,
            "target_page": target_page,
        })

    return _success(
        {
            "total_keywords_received": len(keywords),
            "unique_keywords_classified": len(classified),
            "duplicate_keywords_removed": len(keywords) - len(classified),
            "keywords": classified,
        },
        "Keyword classification completed",
    )


# ==================================================
# WF009: EMPLOYEE TASK ASSIGNMENT
# ==================================================

def assign_employee_task(context: dict) -> dict:
    task = str(
        context.get("task")
        or context.get("task_description")
        or ""
    ).strip()

    if not task:
        return _needs_input(
            "Provide a task description.",
            ["task"],
        )

    priority = str(context.get("priority", "medium")).strip().lower()
    if priority not in {"low", "medium", "high"}:
        raise ValueError("priority must be low, medium, or high")

    deadline = str(context.get("deadline", "")).strip()
    task_lower = task.lower()

    skill_groups = {
        "python": ("python", "programming", "script", "code"),
        "excel": ("excel", "spreadsheet", "workbook"),
        "data": ("data", "analytics", "analysis", "pricing"),
        "reporting": ("report", "reporting"),
        "inventory": ("inventory", "stock", "restock"),
        "marketing": ("marketing", "campaign"),
        "content": ("content", "copywriting"),
        "seo": ("seo", "keyword", "search engine"),
    }

    detected_skills = [
        skill for skill, terms in skill_groups.items()
        if any(term in task_lower for term in terms)
    ]

    supplied_skills = context.get("required_skills", [])
    if isinstance(supplied_skills, str):
        supplied_skills = [
            item.strip().lower()
            for item in supplied_skills.split(",")
            if item.strip()
        ]
    if not isinstance(supplied_skills, list):
        raise ValueError(
            "required_skills must be a list or comma-separated string"
        )

    required_skills = list(dict.fromkeys(
        [str(skill).strip().lower() for skill in supplied_skills
         if str(skill).strip()]
        + detected_skills
    ))

    candidates = []

    for employee in EMPLOYEES:
        employee_skills = {
            skill.strip().lower()
            for skill in str(employee.get("skills", "")).split(",")
            if skill.strip()
        }
        available = employee.get("available", True)
        max_tasks = int(employee.get("max_tasks", 999))
        active_tasks = int(employee.get("active_tasks", 0))

        if not available or active_tasks >= max_tasks:
            continue

        matched = [
            skill for skill in required_skills
            if skill in employee_skills
        ]

        if required_skills and not matched:
            continue

        candidates.append({
            "employee": employee,
            "matched_skills": matched,
            "skill_coverage": (
                len(matched) / len(required_skills)
                if required_skills else 1.0
            ),
        })

    if not candidates:
        return _success({
            "task": task,
            "priority": priority,
            "deadline": deadline or "Not specified",
            "required_skills": required_skills,
            "assignment_status": "escalated",
            "assigned_to": None,
            "escalation_required": True,
            "escalation_reason": (
                "No available employee with matching skills and capacity"
            ),
        }, "No suitable employee available; task escalated")

    selected_record = max(
        candidates,
        key=lambda candidate: (
            candidate["skill_coverage"],
            -int(candidate["employee"].get("active_tasks", 0)),
        ),
    )
    selected = selected_record["employee"]

    return _success({
        "task": task,
        "priority": priority,
        "deadline": deadline or "Not specified",
        "required_skills": required_skills,
        "assignment_status": "recommended",
        "assigned_to": selected["name"],
        "matched_skills": selected_record["matched_skills"],
        "skill_coverage": round(selected_record["skill_coverage"], 2),
        "current_active_tasks": selected["active_tasks"],
        "max_tasks": selected.get("max_tasks", 999),
        "selection_reason": (
            "Best skill coverage among available employees; "
            "workload used as a tie-breaker"
        ),
        "escalation_required": False,
    }, "Task assignment recommendation generated")


# ==================================================
# WF010: WORKFLOW PERFORMANCE REPORT
# ==================================================

def workflow_performance(context: dict) -> dict:
    history = context.get(
        "execution_history",
        EXECUTION_HISTORY,
    )

    if not isinstance(history, list):
        raise ValueError("execution_history must be a list")

    slow_threshold = _number(
        context.get("slow_threshold_seconds", 5),
        "slow_threshold_seconds",
    )

    if slow_threshold < 0:
        raise ValueError(
            "slow_threshold_seconds cannot be negative"
        )

    slow_runs = []
    failures = []
    successes = []
    durations = []
    workflow_counts = {}
    error_counts = {}

    for record in history:
        if not isinstance(record, dict):
            raise ValueError(
                "Each execution history record must be a dictionary"
            )

        duration = _number(
            record.get("duration_seconds", 0),
            "duration_seconds",
        )

        if duration < 0:
            raise ValueError("duration_seconds cannot be negative")

        status = str(record.get("status", "unknown")).lower()
        workflow_id = str(record.get("workflow_id", "unknown"))

        durations.append(duration)
        workflow_counts[workflow_id] = (
            workflow_counts.get(workflow_id, 0) + 1
        )

        if duration > slow_threshold:
            slow_runs.append(record)

        if status in {"success", "completed"}:
            successes.append(record)
        elif status in {"failed", "failure", "error"}:
            failures.append(record)
            error = str(record.get("error", "Unspecified error"))
            error_counts[error] = error_counts.get(error, 0) + 1

    total = len(history)
    success_rate = round(len(successes) / total * 100, 2) if total else 0.0
    failure_rate = round(len(failures) / total * 100, 2) if total else 0.0
    average_duration = round(sum(durations) / total, 2) if total else 0.0

    recommendations = []

    if failures:
        recommendations.append(
            "Investigate failed executions and add retries or input validation where appropriate."
        )
    if slow_runs:
        recommendations.append(
            "Review slow workflows for file I/O, network latency, and unnecessary processing."
        )
    if not failures and not slow_runs:
        recommendations.append(
            "No failures or slow executions detected in the supplied history."
        )

    if total == 0:
        recommendations.append(
            "Run workflows and collect execution history to establish a performance baseline."
        )

    return _success({
        "total_executions": total,
        "successful_executions": len(successes),
        "failed_executions": failures,
        "failure_count": len(failures),
        "success_rate_percent": success_rate,
        "failure_rate_percent": failure_rate,
        "average_duration_seconds": average_duration,
        "slow_executions": slow_runs,
        "slow_execution_count": len(slow_runs),
        "slow_threshold_seconds": slow_threshold,
        "workflow_execution_counts": workflow_counts,
        "frequent_errors": [
            {"error": error, "count": count}
            for error, count in sorted(
                error_counts.items(),
                key=lambda item: item[1],
                reverse=True,
            )
        ],
        "recommendations": recommendations,
    }, "Workflow performance report generated")


# ==================================================
# TOOL REGISTRY
# ==================================================

TOOL_REGISTRY = {
    "WF001": inventory_restock,
    "WF002": validate_prices,
    "WF003": process_vendor_file,
    "WF004": generate_product_content,
    "WF005": get_order_status,
    "WF006": detect_duplicates,
    "WF007": create_campaign_brief,
    "WF008": classify_keywords,
    "WF009": assign_employee_task,
    "WF010": workflow_performance,
}


def execute_tool(
    workflow_id: str,
    context: dict | None = None,
) -> dict:
    """Execute a registered workflow with input and result validation."""
    if not isinstance(workflow_id, str) or workflow_id not in TOOL_REGISTRY:
        raise ValueError(f"Unknown workflow: {workflow_id}")

    if context is None:
        context = {}

    if not isinstance(context, dict):
        raise ValueError("context must be a dictionary")

    result = TOOL_REGISTRY[workflow_id](context)

    if not isinstance(result, dict):
        raise TypeError(
            f"Workflow {workflow_id} returned a non-dictionary result"
        )

    return result

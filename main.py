from pathlib import Path
from typing import BinaryIO
from dotenv import load_dotenv
import openai
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
import os
import re
import json
import sys
import time
import uuid

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "outputs"
PROMPT_PATH = BASE_DIR / "prompts" / "system_prompt.txt"
XLSX_PATH = OUTPUT_DIR / "test_cases.xlsx"


def extract_json_block(text: str) -> dict:
    def parse(candidate: str):
        try:
            return json.loads(candidate)
        except json.JSONDecodeError as e:
            snippet = candidate if len(candidate) <= 2000 else candidate[:2000] + "...(truncated)"
            raise ValueError(
                f"Found a JSON block in the model output but failed to parse it: {e}\n\nRaw block:\n{snippet}"
            ) from e

    marker_pattern = r"JSON_EXPORT_START\s*(\{.*\})\s*JSON_EXPORT_END"
    match = re.search(marker_pattern, text, re.DOTALL)
    if match:
        return parse(match.group(1))

    fence_pattern = r"```json\s*(\{.*?\})\s*```"
    match = re.search(fence_pattern, text, re.DOTALL)
    if match:
        return parse(match.group(1))

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return parse(text[start : end + 1])

    snippet = text if len(text) <= 2000 else text[:2000] + "...(truncated)"
    raise ValueError(f"Could not find any JSON in the model output.\n\nRaw output:\n{snippet}")


def as_lines(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item) for item in value]
    if isinstance(value, dict):
        return [json.dumps(value, ensure_ascii=False)]
    return [str(value)]


EXCEL_FORMULA_TRIGGERS = ("=", "+", "-", "@", "\t", "\r")


def excel_safe(value):
    """Neutralize Excel/CSV formula injection: openpyxl auto-detects a leading
    '=', '+', '-', or '@' as a formula, so untrusted (LLM-generated) cell text
    starting with one of these would execute as a real formula when opened."""
    text = str(value)
    if text and text[0] in EXCEL_FORMULA_TRIGGERS:
        return "'" + text
    return text


def join_for_cell(value):
    return excel_safe("\n".join(as_lines(value)))


def write_header(ws, headers):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def autosize_columns(ws, max_width=45):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        best = 0
        for cell in col_cells:
            value = "" if cell.value is None else str(cell.value)
            best = max(best, max((len(line) for line in value.splitlines()), default=0))
        ws.column_dimensions[col_letter].width = min(max(best + 2, 12), max_width)


def style_sheet(ws):
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def add_kv_sheet(wb, title, rows):
    ws = wb.create_sheet(title)
    write_header(ws, ["Field", "Value"])
    for key, value in rows:
        ws.append([key, join_for_cell(value)])
    style_sheet(ws)
    autosize_columns(ws)
    return ws


def add_list_sheet(wb, title, items, col_name="Item"):
    ws = wb.create_sheet(title)
    write_header(ws, [col_name])
    for item in items:
        ws.append([join_for_cell(item)])
    style_sheet(ws)
    autosize_columns(ws)
    return ws


def normalize_test_case(tc, index):
    if isinstance(tc, dict):
        return {k.strip().lower().replace(" ", "_"): v for k, v in tc.items()}

    if isinstance(tc, str):
        return {
            "id": f"TC-{index:03d}",
            "title": tc,
            "priority": "",
            "type": "",
            "preconditions": [],
            "steps": [],
            "expected_result": [],
            "severity_if_failed": "",
            "notes": "",
        }

    return {
        "id": f"TC-{index:03d}",
        "title": str(tc),
        "priority": "",
        "type": "",
        "preconditions": [],
        "steps": [],
        "expected_result": [],
        "severity_if_failed": "",
        "notes": "",
    }


def build_excel(data: dict, xlsx_path: Path | BinaryIO = XLSX_PATH):
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    add_kv_sheet(
        wb,
        "Summary",
        [
            ("test_objective", data.get("test_objective", "")),
            ("recommendation", data.get("recommendation", "")),
        ],
    )

    scope = data.get("scope", {})
    if isinstance(scope, dict):
        scope_rows = []
        scope_rows.append(("in_scope", scope.get("in_scope", [])))
        scope_rows.append(("out_of_scope", scope.get("out_of_scope", [])))
        add_kv_sheet(wb, "Scope", scope_rows)
    else:
        add_list_sheet(wb, "Scope", scope if scope else [], "Scope Item")

    assumptions = data.get("assumptions", [])
    add_list_sheet(wb, "Assumptions", assumptions, "Assumption")

    risks = data.get("risks", [])
    if isinstance(risks, list) and risks and isinstance(risks[0], dict):
        ws = wb.create_sheet("Risks")
        headers = list(risks[0].keys())
        write_header(ws, headers)
        for risk in risks:
            row = risk if isinstance(risk, dict) else {}
            ws.append([join_for_cell(row.get(h, "")) for h in headers])
        style_sheet(ws)
        autosize_columns(ws)
    else:
        add_list_sheet(wb, "Risks", risks if isinstance(risks, list) else [risks], "Risk")

    clarifications = data.get("clarifications_needed", [])
    add_list_sheet(wb, "Clarifications", clarifications, "Clarification")

    test_cases = data.get("test_cases", [])
    if not isinstance(test_cases, list):
        test_cases = [test_cases]
    ws = wb.create_sheet("Test Cases")
    headers = [
        "ID",
        "Title",
        "Priority",
        "Type",
        "Preconditions",
        "Steps",
        "Expected Result",
        "Severity if Failed",
        "Notes",
    ]
    write_header(ws, headers)

    for i, tc in enumerate(test_cases, start=1):
        tc = normalize_test_case(tc, i)
        ws.append([
            excel_safe(tc.get("id", "")),
            excel_safe(tc.get("title", "")),
            excel_safe(tc.get("priority", "")),
            excel_safe(tc.get("type", "")),
            join_for_cell(tc.get("preconditions", [])),
            join_for_cell(tc.get("steps", [])),
            join_for_cell(tc.get("expected_result", [])),
            excel_safe(tc.get("severity_if_failed", "")),
            excel_safe(tc.get("notes", "")),
        ])

    style_sheet(ws)
    autosize_columns(ws)

    exploratory = data.get("exploratory_charters", [])
    if isinstance(exploratory, list) and exploratory and isinstance(exploratory[0], dict):
        ws = wb.create_sheet("Exploratory")
        headers = list(exploratory[0].keys())
        write_header(ws, headers)
        for charter in exploratory:
            row = charter if isinstance(charter, dict) else {}
            ws.append([join_for_cell(row.get(h, "")) for h in headers])
        style_sheet(ws)
        autosize_columns(ws)
    else:
        add_list_sheet(wb, "Exploratory", exploratory if isinstance(exploratory, list) else [exploratory], "Charter")

    coverage_gaps = data.get("coverage_gaps", [])
    add_list_sheet(wb, "Coverage Gaps", coverage_gaps, "Gap")

    if isinstance(xlsx_path, Path):
        xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)


MAX_JSON_RETRIES = 2

DEMO_LIMIT_MESSAGE = (
    "Demo usage limit reached for this month. Please contact me for a live "
    "walkthrough: hadeelbeiruti@gmail.com"
)


def _create_response(client, **kwargs):
    try:
        return client.responses.create(**kwargs)
    except openai.RateLimitError as e:
        raise ValueError(DEMO_LIMIT_MESSAGE) from e


def generate_qa_package(story_text: str, on_retry=None) -> tuple[str, dict, dict]:
    load_dotenv()

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OPENAI_API_KEY is missing in .env")

    client = OpenAI(api_key=api_key)
    system_prompt = PROMPT_PATH.read_text(encoding="utf-8")
    model = os.getenv("OPENAI_MODEL", "gpt-5.5")

    start = time.perf_counter()
    response = _create_response(
        client,
        model=model,
        instructions=system_prompt,
        input=story_text,
    )
    text = response.output_text
    print(f"[generate_qa_package] initial call took {time.perf_counter() - start:.1f}s", flush=True)

    for attempt in range(MAX_JSON_RETRIES + 1):
        try:
            data = extract_json_block(text)
            break
        except ValueError as e:
            if attempt == MAX_JSON_RETRIES:
                raise
            print(f"[generate_qa_package] JSON parse failed (attempt {attempt + 1}/{MAX_JSON_RETRIES + 1}): {e}", flush=True)
            if on_retry:
                on_retry(attempt + 1, MAX_JSON_RETRIES + 1)
            response = _create_response(
                client,
                model=model,
                previous_response_id=response.id,
                input=(
                    f"Your previous response's JSON export could not be parsed: {e}\n\n"
                    "Resend your full response again (human-readable report and JSON export), "
                    "with the JSON export corrected so it is valid JSON matching the required "
                    "schema, placed exactly between the JSON_EXPORT_START and JSON_EXPORT_END markers."
                ),
            )
            text = response.output_text

    print(f"[generate_qa_package] total time including retries: {time.perf_counter() - start:.1f}s", flush=True)

    run_dir = OUTPUT_DIR / f"run_{uuid.uuid4().hex[:8]}"
    run_dir.mkdir(parents=True, exist_ok=True)
    txt_path = run_dir / "latest_output.txt"
    json_path = run_dir / "latest_output.json"
    xlsx_path = run_dir / "test_cases.xlsx"
    story_path = run_dir / "story.txt"

    txt_path.write_text(text, encoding="utf-8")
    json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    story_path.write_text(story_text, encoding="utf-8")
    build_excel(data, xlsx_path)

    paths = {"txt": txt_path, "json": json_path, "xlsx": xlsx_path, "story": story_path}
    return text, data, paths


def main():
    story_path = BASE_DIR / "story.txt"
    story = story_path.read_text(encoding="utf-8")

    print("Sending request to OpenAI...", flush=True)
    try:
        text, _, paths = generate_qa_package(story)
        print("\n" + "=" * 80)
        print("QA AGENT OUTPUT")
        print("=" * 80 + "\n")
        print(text)
        print("\nSaved JSON  ->", paths["json"])
        print("Saved Excel ->", paths["xlsx"])
        print("Saved Text  ->", paths["txt"])
    except Exception as e:
        print(f"\nRequest failed: {e}", file=sys.stderr)
        raise


if __name__ == "__main__":
    main()
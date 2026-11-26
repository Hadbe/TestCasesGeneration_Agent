import json
import os
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

import azure_devops
from main import OUTPUT_DIR, as_lines, build_excel, generate_qa_package

load_dotenv()

MAX_PARALLEL_GENERATIONS = 4

TEST_CASE_COLUMNS = [
    "id", "title", "priority", "type",
    "preconditions", "steps", "expected_result",
    "severity_if_failed", "notes",
]

EXAMPLE_STORY = """User Story:
As an authenticated user, I want to change my email address and verify it with a one time code so that my account remains secure.

Acceptance Criteria:
1. The user can enter a new email address.
2. The system validates that the email format is valid.
3. The system sends a verification code to the new email address.
4. The email address is updated only after the correct code is entered.
5. The verification code expires after a defined time.
6. The user cannot use an email address that is already taken.
7. The user can resend the verification code.
8. The user can cancel the change before verification is completed.
9. If the code is incorrect or expired, the change must not be saved.
10. The updated email should remain after refresh.
"""

st.set_page_config(page_title="QA Analysis & Test Case Generator", layout="wide")

SESSION_RUNS_PATH = OUTPUT_DIR / "session_runs.json"


def get_session_id():
    """A per-visitor id persisted in the URL's query string, so Run History
    survives a page refresh (session_state alone does not: a refresh starts
    a brand new Streamlit session)."""
    sid = st.query_params.get("sid")
    if not sid:
        sid = uuid.uuid4().hex[:16]
        st.query_params["sid"] = sid
    return sid


SESSION_ID = get_session_id()


def _load_session_runs():
    if not SESSION_RUNS_PATH.exists():
        return {}
    try:
        return json.loads(SESSION_RUNS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _save_session_runs(data):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SESSION_RUNS_PATH.write_text(json.dumps(data), encoding="utf-8")


def test_case_to_row(tc):
    return {
        "id": tc.get("id", ""),
        "title": tc.get("title", ""),
        "priority": tc.get("priority", ""),
        "type": tc.get("type", ""),
        "preconditions": "\n".join(as_lines(tc.get("preconditions"))),
        "steps": "\n".join(as_lines(tc.get("steps"))),
        "expected_result": "\n".join(as_lines(tc.get("expected_result"))),
        "severity_if_failed": tc.get("severity_if_failed", ""),
        "notes": tc.get("notes", ""),
    }


def row_to_test_case(row):
    return {
        "id": row.get("id", ""),
        "title": row.get("title", ""),
        "priority": row.get("priority", ""),
        "type": row.get("type", ""),
        "preconditions": [line for line in str(row.get("preconditions", "")).split("\n") if line.strip()],
        "steps": [line for line in str(row.get("steps", "")).split("\n") if line.strip()],
        "expected_result": [line for line in str(row.get("expected_result", "")).split("\n") if line.strip()],
        "severity_if_failed": row.get("severity_if_failed", ""),
        "notes": row.get("notes", ""),
    }


def escape_markdown_link_text(text):
    """Prevent LLM-generated text from breaking out of markdown link brackets
    (e.g. a title containing '](http://evil.com)' could hijack the link's
    actual destination) by escaping markdown-significant characters."""
    text = str(text)
    for ch in ("\\", "[", "]", "(", ")"):
        text = text.replace(ch, "\\" + ch)
    return text


def story_title(text):
    for line in text.strip().splitlines():
        candidate = line.strip()
        if candidate.lower().startswith("user story:"):
            candidate = candidate[len("user story:"):].strip()
        if candidate:
            return candidate[:80] + ("…" if len(candidate) > 80 else "")
    return ""


def render_result(result, key_prefix):
    display_text = result["text"]
    if result.get("story"):
        title = story_title(result["story"])
    else:
        title = result["data"].get("test_objective", "")
    if title:
        display_text = f"{title}\n\n{display_text}"

    st.text_area("Executive / QA Output", value=display_text, height=300, key=f"{key_prefix}_report")

    st.markdown("**Test Cases — review and edit before export**")
    st.caption("Edit cells, add/delete rows as needed. Downloads below reflect your edits.")

    df = pd.DataFrame(
        [test_case_to_row(tc) for tc in result["data"].get("test_cases", [])],
        columns=TEST_CASE_COLUMNS,
    )
    edited_df = st.data_editor(df, num_rows="dynamic", width="stretch", key=f"{key_prefix}_editor")

    with st.expander("Other sections (read-only)"):
        st.json({k: v for k, v in result["data"].items() if k != "test_cases"})

    edited_test_cases = [row_to_test_case(row) for row in edited_df.to_dict("records")]
    export_data = {**result["data"], "test_cases": edited_test_cases}

    json_bytes = json.dumps(export_data, ensure_ascii=False, indent=2).encode("utf-8")
    xlsx_buffer = BytesIO()
    build_excel(export_data, xlsx_buffer)

    paths = result.get("paths", {})
    if paths.get("json"):
        paths["json"].write_bytes(json_bytes)
    if paths.get("xlsx"):
        paths["xlsx"].write_bytes(xlsx_buffer.getvalue())

    col_a, col_b = st.columns(2)
    col_a.download_button(
        "Download JSON",
        data=json_bytes,
        file_name="latest_output.json",
        mime="application/json",
        key=f"{key_prefix}_dl_json",
    )
    col_b.download_button(
        "Download Excel",
        data=xlsx_buffer.getvalue(),
        file_name="test_cases.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{key_prefix}_dl_xlsx",
    )

    st.markdown("**Push to Azure DevOps**")
    with st.expander("Configure and push test cases to Azure DevOps", expanded=False):
        col_org, col_proj = st.columns(2)
        ado_org = col_org.text_input(
            "Organization", value=os.getenv("AZURE_DEVOPS_ORG", ""), key=f"{key_prefix}_ado_org"
        )
        ado_project = col_proj.text_input(
            "Project", value=os.getenv("AZURE_DEVOPS_PROJECT", ""), key=f"{key_prefix}_ado_project"
        )
        ado_pat = st.text_input(
            "Personal Access Token",
            value=os.getenv("AZURE_DEVOPS_PAT", ""),
            type="password",
            key=f"{key_prefix}_ado_pat",
        )

        link_mode = st.radio(
            "Link test cases to",
            ["Create a new User Story", "Link to an existing User Story"],
            key=f"{key_prefix}_ado_link_mode",
            horizontal=True,
        )

        ado_story_id = None
        ado_story_title = None
        if link_mode == "Link to an existing User Story":
            ado_story_id = st.number_input(
                "Existing User Story ID", min_value=1, step=1, value=None,
                placeholder="e.g. 4521", key=f"{key_prefix}_ado_story_id"
            )
        else:
            ado_story_title = st.text_input(
                "New User Story title", value=title or "New User Story", key=f"{key_prefix}_ado_story_title"
            )

        if st.button("Push to Azure DevOps", key=f"{key_prefix}_ado_push"):
            if not (ado_org and ado_project and ado_pat):
                st.error("Organization, Project, and Personal Access Token are all required.")
            elif not edited_test_cases:
                st.error("There are no test cases to push.")
            elif link_mode == "Link to an existing User Story" and not ado_story_id:
                st.error("Please enter the existing User Story ID to link to.")
            else:
                with st.spinner("Pushing to Azure DevOps..."):
                    try:
                        created, story_info = azure_devops.push_test_cases(
                            ado_org, ado_project, ado_pat, edited_test_cases,
                            user_story_id=int(ado_story_id) if ado_story_id else None,
                            user_story_title=ado_story_title if link_mode == "Create a new User Story" else None,
                        )
                        st.success(f"Pushed {len(created)} test case(s) to Azure DevOps.")
                        if story_info:
                            safe_title = escape_markdown_link_text(story_info['title'])
                            st.markdown(f"[User Story #{story_info['id']}: {safe_title}]({story_info['url']})")
                        for c in created:
                            safe_title = escape_markdown_link_text(c['title'])
                            st.markdown(f"- [#{c['id']}: {safe_title}]({c['url']})")
                    except azure_devops.AzureDevOpsError as e:
                        st.error(f"Azure DevOps push failed: {e}")


def track_run(paths):
    """Remember this run as belonging to the current visitor (identified by
    SESSION_ID), so Run History only ever shows runs this visitor generated,
    not every visitor's runs on the shared server. Persisted to disk (keyed
    by SESSION_ID, which lives in the URL) so it survives a page refresh."""
    run_dir_name = paths["json"].parent.name
    st.session_state.setdefault("my_run_ids", set()).add(run_dir_name)
    data = _load_session_runs()
    owned = data.setdefault(SESSION_ID, [])
    if run_dir_name not in owned:
        owned.append(run_dir_name)
    _save_session_runs(data)


def list_past_runs():
    if not OUTPUT_DIR.exists():
        return []
    my_run_ids = set(st.session_state.get("my_run_ids", set())) | set(
        _load_session_runs().get(SESSION_ID, [])
    )
    runs = []
    for run_dir in sorted(OUTPUT_DIR.glob("run_*"), key=lambda p: p.stat().st_mtime, reverse=True):
        if run_dir.name not in my_run_ids:
            continue
        json_path = run_dir / "latest_output.json"
        if not json_path.exists():
            continue
        runs.append({
            "dir": run_dir,
            "mtime": run_dir.stat().st_mtime,
            "json": json_path,
            "txt": run_dir / "latest_output.txt",
            "xlsx": run_dir / "test_cases.xlsx",
            "story": run_dir / "story.txt",
        })
    return runs


with st.sidebar:
    st.header("Run History")
    st.caption(
        "Only visible to you, for this visit's link. It may be cleared if the "
        "demo restarts after a long period of inactivity."
    )
    past_runs = list_past_runs()
    if not past_runs:
        st.caption("No past runs yet.")
    for run in past_runs:
        try:
            objective = json.loads(run["json"].read_text(encoding="utf-8")).get("test_objective", "")
        except (json.JSONDecodeError, OSError):
            objective = ""
        timestamp = datetime.fromtimestamp(run["mtime"]).strftime("%Y-%m-%d %H:%M")
        caption = objective[:60] + ("…" if len(objective) > 60 else "")
        clicked = st.button(timestamp, key=f"history_{run['dir'].name}", width="stretch")
        if caption:
            st.caption(caption)
        if clicked:
            st.session_state["history_result"] = {
                "text": run["txt"].read_text(encoding="utf-8") if run["txt"].exists() else "",
                "data": json.loads(run["json"].read_text(encoding="utf-8")),
                "story": run["story"].read_text(encoding="utf-8") if run["story"].exists() else None,
                "paths": {"txt": run["txt"], "json": run["json"], "xlsx": run["xlsx"]},
            }

st.title("QA Analysis & Test Case Generator")
st.write("Paste a user story, generate QA analysis, and download the outputs.")

history_result = st.session_state.get("history_result")
if history_result:
    st.divider()
    st.subheader("Viewing past run")
    render_result(history_result, key_prefix=f"history_{history_result['paths']['json']}")
    st.divider()

single_tab, batch_tab = st.tabs(["Single Story", "Batch — Multiple Stories"])

with single_tab:

    def load_example_callback():
        st.session_state["story"] = EXAMPLE_STORY

    story = st.text_area(
        "User Story",
        height=300,
        placeholder="Paste your user story and acceptance criteria here...",
        key="story",
    )

    col1, col2 = st.columns([1, 1])
    generate = col1.button("Generate QA Package")
    col2.button("Load Example", on_click=load_example_callback)

    if generate:
        if not story.strip():
            st.error("Please paste a user story first.")
        else:
            with st.status("Generating QA package...", expanded=False) as status:
                def on_retry(attempt, total, status=status):
                    status.update(label=f"Model's JSON needed a fix — retrying ({attempt}/{total - 1})...")

                try:
                    text, data, paths = generate_qa_package(story, on_retry=on_retry)
                    st.session_state["result"] = {"text": text, "data": data, "paths": paths, "story": story}
                    track_run(paths)
                    status.update(label="Done", state="complete")
                except Exception as e:
                    st.session_state["result"] = None
                    status.update(label="Failed", state="error")
                    st.error(f"Generation failed: {e}")

    result = st.session_state.get("result")
    if result:
        st.success("Done")
        render_result(result, key_prefix=f"single_{result['paths']['json']}")

with batch_tab:
    st.write("Queue several user stories, then generate QA packages for all of them in one pass.")

    def make_label(text, index):
        first_line = next((line.strip() for line in text.strip().splitlines() if line.strip()), "")
        first_line = first_line.replace("User Story:", "").strip()
        if not first_line:
            first_line = f"Story {index + 1}"
        return first_line[:50] + ("…" if len(first_line) > 50 else "")

    def add_to_queue_callback():
        text = st.session_state.get("batch_new_story", "").strip()
        if not text:
            return
        queue = st.session_state.setdefault("story_queue", [])
        queue.append({
            "id": str(uuid.uuid4()),
            "text": text,
            "label": make_label(text, len(queue)),
            "status": "queued",
            "result": None,
            "error": None,
        })
        st.session_state["batch_new_story"] = ""

    def remove_from_queue_callback(story_id):
        st.session_state["story_queue"] = [
            s for s in st.session_state.get("story_queue", []) if s["id"] != story_id
        ]

    st.text_area(
        "Add a story to the queue",
        height=150,
        placeholder="Paste a user story and acceptance criteria here, then click 'Add to Queue'...",
        key="batch_new_story",
    )
    st.button("Add to Queue", on_click=add_to_queue_callback)

    queue = st.session_state.get("story_queue", [])
    st.subheader(f"Queue ({len(queue)})")

    if not queue:
        st.caption("No stories queued yet.")
    else:
        pending = [s for s in queue if s["status"] != "done"]
        if st.button(f"Generate All ({len(pending)} pending)", disabled=not pending):
            progress = st.progress(0.0, text=f"Generating {len(pending)} stories in parallel...")
            workers = min(MAX_PARALLEL_GENERATIONS, len(pending))
            with ThreadPoolExecutor(max_workers=workers) as executor:
                future_to_item = {
                    executor.submit(generate_qa_package, item["text"]): item for item in pending
                }
                completed = 0
                for future in as_completed(future_to_item):
                    item = future_to_item[future]
                    completed += 1
                    try:
                        text, data, paths = future.result()
                        item["result"] = {"text": text, "data": data, "paths": paths, "story": item["text"]}
                        track_run(paths)
                        item["status"] = "done"
                        item["error"] = None
                    except Exception as e:
                        item["status"] = "error"
                        item["error"] = str(e)
                    progress.progress(
                        completed / len(pending),
                        text=f"Completed {completed}/{len(pending)}: {item['label']}",
                    )
            progress.progress(1.0, text="Done")

        status_icon = {"done": "✅", "error": "⚠️", "queued": "⏳"}
        tab_labels = [f"{status_icon[s['status']]} {s['label']}" for s in queue]
        story_tabs = st.tabs(tab_labels)

        for tab, item in zip(story_tabs, queue):
            with tab:
                if st.button("Remove from queue", key=f"remove_{item['id']}"):
                    remove_from_queue_callback(item["id"])
                    st.rerun()

                if item["status"] == "error":
                    st.error(f"Generation failed: {item['error']}")
                    if st.button("Retry", key=f"retry_{item['id']}"):
                        with st.status("Retrying...", expanded=False) as status:
                            def on_retry(attempt, total, status=status):
                                status.update(label=f"Model's JSON needed a fix — retrying ({attempt}/{total - 1})...")

                            try:
                                text, data, paths = generate_qa_package(item["text"], on_retry=on_retry)
                                item["result"] = {"text": text, "data": data, "paths": paths, "story": item["text"]}
                                track_run(paths)
                                item["status"] = "done"
                                item["error"] = None
                                status.update(label="Done", state="complete")
                            except Exception as e:
                                item["error"] = str(e)
                                status.update(label="Failed", state="error")
                        st.rerun()
                elif item["status"] == "done":
                    render_result(item["result"], key_prefix=item["id"])
                else:
                    st.caption("Not generated yet — click 'Generate All' above.")
                    st.text_area(
                        "Story preview",
                        value=item["text"],
                        height=150,
                        disabled=True,
                        key=f"preview_{item['id']}",
                    )

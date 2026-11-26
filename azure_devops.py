import base64
from urllib.parse import quote
from xml.sax.saxutils import escape

import requests

API_VERSION = "7.1"

PRIORITY_MAP = {
    "critical": 1,
    "high": 2,
    "medium": 3,
    "low": 4,
}


class AzureDevOpsError(Exception):
    pass


def _auth_header(pat):
    token = base64.b64encode(f":{pat}".encode("utf-8")).decode("ascii")
    return {
        "Authorization": f"Basic {token}",
        "Content-Type": "application/json-patch+json",
    }


def _base_url(org, project):
    return f"https://dev.azure.com/{quote(org, safe='')}/{quote(project, safe='')}/_apis/wit"


def _priority_value(priority_text):
    return PRIORITY_MAP.get(str(priority_text or "").strip().lower(), 3)


def _steps_xml(steps, expected_results):
    if isinstance(steps, str):
        steps = [steps]
    if isinstance(expected_results, str):
        expected_results = [expected_results]
    steps = list(steps) if steps else [""]
    expected_results = list(expected_results) if expected_results else []
    rows = []
    for i, action in enumerate(steps):
        expected = expected_results[i] if i < len(expected_results) else ""
        rows.append(
            f'<step id="{i + 1}" type="ActionStep">'
            f'<parameterizedString isformatted="true">{escape(str(action))}</parameterizedString>'
            f'<parameterizedString isformatted="true">{escape(str(expected))}</parameterizedString>'
            f"<description/></step>"
        )
    return f'<steps id="0" last="{len(rows)}">' + "".join(rows) + "</steps>"


def _request(method, url, pat, json_body=None):
    try:
        resp = requests.request(method, url, headers=_auth_header(pat), json=json_body, timeout=30)
    except requests.RequestException as e:
        raise AzureDevOpsError(f"Could not reach Azure DevOps: {e}") from e

    if not resp.ok:
        detail = resp.text[:500]
        raise AzureDevOpsError(f"Azure DevOps API returned {resp.status_code}: {detail}")

    try:
        return resp.json()
    except ValueError as e:
        snippet = resp.text[:500] if resp.text else "(empty response body)"
        raise AzureDevOpsError(
            f"Azure DevOps returned a non-JSON response (HTTP {resp.status_code}). This usually means "
            "the Organization/Project name is wrong, or the Personal Access Token is invalid, expired, "
            f"or missing the 'Work Items (Read & Write)' scope.\n\nResponse snippet:\n{snippet}"
        ) from e


def create_work_item(org, project, pat, work_item_type, fields, relations=None):
    url = f"{_base_url(org, project)}/workitems/${quote(work_item_type)}?api-version={API_VERSION}"
    patch = [{"op": "add", "path": f"/fields/{name}", "value": value} for name, value in fields.items()]
    for relation in relations or []:
        patch.append({"op": "add", "path": "/relations/-", "value": relation})
    return _request("POST", url, pat, patch)


def get_work_item(org, project, pat, work_item_id):
    url = f"{_base_url(org, project)}/workitems/{work_item_id}?api-version={API_VERSION}"
    return _request("GET", url, pat)


def push_test_cases(org, project, pat, test_cases, user_story_id=None, user_story_title=None,
                     test_case_type="Test Case", user_story_type="User Story"):
    """Create Test Case work items in Azure DevOps, linked to a User Story.

    If user_story_id is given, links to that existing work item.
    Otherwise, if user_story_title is given, creates a new User Story first.
    Returns (created_test_cases, user_story_info) where each item has id/title/url.
    """
    user_story_info = None
    parent_url = None

    if user_story_id:
        parent = get_work_item(org, project, pat, user_story_id)
        parent_url = parent["url"]
        user_story_info = {
            "id": parent["id"],
            "title": parent["fields"].get("System.Title", ""),
            "url": parent["_links"]["html"]["href"],
        }
    elif user_story_title:
        parent = create_work_item(org, project, pat, user_story_type, {"System.Title": user_story_title})
        parent_url = parent["url"]
        user_story_info = {
            "id": parent["id"],
            "title": parent["fields"].get("System.Title", ""),
            "url": parent["_links"]["html"]["href"],
        }

    created = []
    for tc in test_cases:
        fields = {
            "System.Title": tc.get("title") or "Untitled test case",
            "Microsoft.VSTS.TCM.Steps": _steps_xml(tc.get("steps"), tc.get("expected_result")),
            "Microsoft.VSTS.Common.Priority": _priority_value(tc.get("priority")),
        }
        notes = tc.get("notes")
        preconditions = tc.get("preconditions")
        description_parts = []
        if preconditions:
            items = preconditions if isinstance(preconditions, list) else [preconditions]
            description_parts.append("Preconditions: " + "; ".join(str(p) for p in items))
        if notes:
            description_parts.append(str(notes))
        if description_parts:
            fields["System.Description"] = escape(" | ".join(description_parts))

        relations = None
        if parent_url:
            relations = [{"rel": "Microsoft.VSTS.Common.TestedBy-Reverse", "url": parent_url}]

        item = create_work_item(org, project, pat, test_case_type, fields, relations)
        created.append({
            "id": item["id"],
            "title": item["fields"].get("System.Title", ""),
            "url": item["_links"]["html"]["href"],
        })

    return created, user_story_info

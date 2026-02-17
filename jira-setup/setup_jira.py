#!/usr/bin/env python3
"""
setup_jira.py
=============
Creates the AccessVBA → Azure SQL Migration Jira project, epics, and
user stories from scratch.

Usage:
    1. Copy .env.example to .env and fill in your credentials
    2. pip install requests python-dotenv
    3. python setup_jira.py

The script is idempotent for epics and stories — it checks whether an
issue with the same summary already exists before creating it, so it is
safe to re-run if it fails partway through.
"""

import os
import sys
import json
import time
import requests
from requests.auth import HTTPBasicAuth

# ---------------------------------------------------------------------------
# Load .env
# ---------------------------------------------------------------------------
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # .env values must already be in the environment

JIRA_BASE_URL = os.environ.get("JIRA_BASE_URL", "").rstrip("/")
JIRA_EMAIL    = os.environ.get("JIRA_EMAIL", "")
JIRA_API_TOKEN = os.environ.get("JIRA_API_TOKEN", "")

if not all([JIRA_BASE_URL, JIRA_EMAIL, JIRA_API_TOKEN]):
    sys.exit("ERROR: Set JIRA_BASE_URL, JIRA_EMAIL, and JIRA_API_TOKEN in .env")

AUTH    = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)
HEADERS = {"Accept": "application/json", "Content-Type": "application/json"}

# ---------------------------------------------------------------------------
# Project definition
# ---------------------------------------------------------------------------
PROJECT_NAME = "AccessVBA to Azure SQL Migration"
PROJECT_KEY  = "AVBA"          # change if already taken
PROJECT_DESC = (
    "Migrate the Q-1019 Access database backend from a shared-drive .accdb "
    "to Azure SQL (AAD auth), and move frontend distribution to SharePoint."
)

# ---------------------------------------------------------------------------
# Board content
# Epics → list of (summary, description, stories)
# Each story is (summary, acceptance_criteria_lines)
# ---------------------------------------------------------------------------
BOARD = [
    {
        "epic_summary": "Azure SQL Setup & Schema",
        "epic_description": (
            "Provision the Azure SQL instance, migrate the Access schema using SSMA, "
            "configure Azure AD authentication, and validate network access."
        ),
        "stories": [
            (
                "Provision Azure SQL Database instance",
                [
                    "Azure SQL serverless instance created in the correct subscription and resource group",
                    "Firewall rules allow office IP range and VPN exit IPs",
                    "Connection confirmed from a developer machine via SSMS",
                    "Service tier documented in project README",
                ],
            ),
            (
                "Migrate Access schema to Azure SQL using SSMA",
                [
                    "All user tables present with correct column names, types, and nullability",
                    "All indexes migrated, including composite indexes on SalesOrders",
                    "Foreign key relationships enforced in Azure SQL",
                    "SSMA migration report reviewed and zero critical errors",
                ],
            ),
            (
                "Configure Azure AD authentication for application login",
                [
                    "AAD group created for Q-1019 database users",
                    "Application login granted minimum required permissions (no db_owner)",
                    "At least two test users can authenticate via AAD and open the database",
                    "No SQL passwords stored in VBA or config files",
                ],
            ),
            (
                "Validate network access for office and remote users",
                [
                    "Office users can reach Azure SQL directly without VPN",
                    "Remote users can reach Azure SQL via VPN with latency under 500 ms",
                    "basRemoteAccess.TestBackendConnection() returns True from both environments",
                ],
            ),
        ],
    },
    {
        "epic_summary": "Data Migration",
        "epic_description": (
            "Move all existing backend data to Azure SQL, validate integrity, "
            "and perform the final cutover delta sync."
        ),
        "stories": [
            (
                "Run initial data load from Access BE to Azure SQL",
                [
                    "SSMA data migration completed for all tables",
                    "Row counts match between Access BE and Azure SQL for every table",
                    "No truncation errors logged during migration",
                ],
            ),
            (
                "Validate order numbering data integrity after migration",
                [
                    "OrderSeq table values match the Access BE exactly",
                    "No gaps or duplicates in SalesOrders.OrderNumber",
                    "Highest migrated order number is within the correct year band (576xxx)",
                ],
            ),
            (
                "Validate reference and lookup table data",
                [
                    "All lookup/reference tables are present and row counts match",
                    "Spot-check of 10 records per lookup table confirms values are correct",
                ],
            ),
            (
                "Perform cutover delta sync and lock old backend",
                [
                    "Any records created in Access BE after initial load are synced to Azure SQL",
                    "Old .accdb BE file renamed to .bak to prevent further writes",
                    "Azure SQL confirmed as sole source of truth before FE re-link",
                ],
            ),
        ],
    },
    {
        "epic_summary": "FE Re-link & VBA Updates",
        "epic_description": (
            "Update all VBA modules to use ODBC connection strings instead of file paths, "
            "fix Access-specific SQL, and adapt sequence allocation for SQL Server."
        ),
        "stories": [
            (
                "Update basRelinkTables to use ODBC connection strings",
                [
                    "RelinkAllTables() builds a DSN-less ODBC connection string from tblConfig keys",
                    "All linked tables successfully re-point to Azure SQL after running RelinkAllTables()",
                    "GetLinkedBackends() diagnostic confirms all tables share the same ODBC connection string",
                    "No hard-coded file paths remain in basRelinkTables",
                ],
            ),
            (
                "Update basRemoteAccess to test ODBC connectivity",
                [
                    "TestBackendConnection() opens an ODBC connection to Azure SQL (not a file path)",
                    "Returns False with a clear message when Azure SQL is unreachable",
                    "Latency measurement still works and is logged",
                ],
            ),
            (
                "Update tblConfig schema for Azure SQL connection parameters",
                [
                    "tblConfig stores AzureSQLServer, AzureSQLDatabase, and AzureSQLAuthMode keys",
                    "BackendPath key retired or repurposed as a legacy fallback comment",
                    "basConfig updated with new default keys; no credentials stored in tblConfig",
                ],
            ),
            (
                "Update basSeqAllocator for SQL Server concurrency",
                [
                    "Sequence allocation uses a transaction with appropriate isolation level",
                    "No duplicate order numbers generated under simulated concurrent load (5 users)",
                    "Retry logic still functions if a transaction conflict occurs",
                ],
            ),
            (
                "Remediate Access-specific SQL in all saved queries",
                [
                    "All saved queries audited for Access-only syntax (IIf, *, Format, TRANSFORM)",
                    "Incompatible queries rewritten as passthrough queries or T-SQL views",
                    "All forms that use saved queries open without error against Azure SQL backend",
                ],
            ),
        ],
    },
    {
        "epic_summary": "FE Distribution via SharePoint",
        "epic_description": (
            "Set up SharePoint as the distribution channel for the Access frontend, "
            "update basVersion for cloud-based update detection, and roll out to all users."
        ),
        "stories": [
            (
                "Set up SharePoint document library for FE template distribution",
                [
                    "SharePoint library created with versioning enabled",
                    "Q1019_FE_TEMPLATE.accdb stored in library with correct permissions",
                    "Only admins can upload new versions; all users can download",
                ],
            ),
            (
                "Update basVersion to detect newer FE version from SharePoint",
                [
                    "On startup the FE checks the SharePoint library for a version newer than the local copy",
                    "User is prompted to download and replace their local copy if a newer version exists",
                    "Version check fails gracefully (no crash) if SharePoint is unreachable",
                ],
            ),
            (
                "Write user installation and update guide",
                [
                    "Step-by-step guide covers first-time install: download from SharePoint, save locally, open",
                    "Guide covers the update prompt and how to accept/defer",
                    "Guide lists ODBC Driver 18 for SQL Server as a prerequisite with download link",
                    "Guide reviewed and approved by at least one non-technical user",
                ],
            ),
            (
                "Pilot rollout with 2-3 users and resolve issues",
                [
                    "At least two pilot users (one office, one remote) successfully open the FE from SharePoint",
                    "End-to-end order creation tested: new order saved, order number in 576xxx range",
                    "All pilot issues logged and resolved before full rollout",
                ],
            ),
            (
                "Full user rollout and decommission shared-drive FE copies",
                [
                    "All 6-15 users migrated to SharePoint-distributed FE",
                    "Old shared-drive FE copies removed or renamed to prevent accidental use",
                    "Confirmation from each user that they can open and use the new FE",
                ],
            ),
        ],
    },
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def jira_get(path, params=None):
    url = f"{JIRA_BASE_URL}/rest/api/3{path}"
    r = requests.get(url, auth=AUTH, headers=HEADERS, params=params, timeout=15)
    r.raise_for_status()
    return r.json()


def jira_post(path, body):
    url = f"{JIRA_BASE_URL}/rest/api/3{path}"
    r = requests.post(url, auth=AUTH, headers=HEADERS,
                      data=json.dumps(body), timeout=15)
    if not r.ok:
        print(f"  ERROR {r.status_code}: {r.text[:400]}")
        r.raise_for_status()
    return r.json()


def search_issues(jql, fields="summary,issuetype"):
    body = {"jql": jql, "fields": fields.split(","), "maxResults": 50}
    result = jira_post("/search/jql", body)
    return result.get("issues", [])


def issue_exists(project_key, summary, issue_type):
    safe = summary.replace('"', '\\"')
    jql  = f'project = "{project_key}" AND issuetype = "{issue_type}" AND summary ~ "{safe}"'
    hits = search_issues(jql)
    # exact match check (~ is contains, not equals)
    for h in hits:
        if h["fields"]["summary"].strip() == summary.strip():
            return h["key"]
    return None


def get_account_id():
    me = jira_get("/myself")
    return me["accountId"]


def adf_doc(text):
    """Wrap plain text as an Atlassian Document Format doc."""
    return {
        "type": "doc",
        "version": 1,
        "content": [
            {
                "type": "paragraph",
                "content": [{"type": "text", "text": text}],
            }
        ],
    }


def adf_acceptance_criteria(lines):
    """Build an ADF doc with an AC heading and a bullet list."""
    items = [
        {
            "type": "listItem",
            "content": [
                {
                    "type": "paragraph",
                    "content": [{"type": "text", "text": line}],
                }
            ],
        }
        for line in lines
    ]
    return {
        "type": "doc",
        "version": 1,
        "content": [
            {
                "type": "heading",
                "attrs": {"level": 3},
                "content": [{"type": "text", "text": "Acceptance Criteria"}],
            },
            {"type": "bulletList", "content": items},
        ],
    }


# ---------------------------------------------------------------------------
# Project creation
# ---------------------------------------------------------------------------

def get_or_create_project():
    # Check if project key already exists
    try:
        proj = jira_get(f"/project/{PROJECT_KEY}")
        print(f"  Project {PROJECT_KEY} already exists: {proj['name']}")
        return proj["key"], proj["id"]
    except requests.HTTPError as e:
        if e.response.status_code != 404:
            raise

    print(f"  Creating project {PROJECT_KEY} ...")
    account_id = get_account_id()

    body = {
        "key":         PROJECT_KEY,
        "name":        PROJECT_NAME,
        "description": PROJECT_DESC,
        "projectTypeKey":     "software",
        "projectTemplateKey": "com.pyxis.greenhopper.jira:gh-simplified-kanban-classic",
        "leadAccountId":      account_id,
        "assigneeType":       "UNASSIGNED",
    }

    result = jira_post("/project", body)
    print(f"  Created: {result['key']} (id={result['id']})")
    return result["key"], str(result["id"])


# ---------------------------------------------------------------------------
# Epic + story creation
# ---------------------------------------------------------------------------

def get_or_create_epic(project_key, summary, description):
    existing = issue_exists(project_key, summary, "Epic")
    if existing:
        print(f"    Epic already exists: {existing}  \"{summary}\"")
        return existing

    body = {
        "fields": {
            "project":     {"key": project_key},
            "summary":     summary,
            "description": adf_doc(description),
            "issuetype":   {"name": "Epic"},
        }
    }
    result = jira_post("/issue", body)
    key = result["key"]
    print(f"    Created epic: {key}  \"{summary}\"")
    time.sleep(0.3)
    return key


def get_or_create_story(project_key, epic_key, summary, ac_lines):
    existing = issue_exists(project_key, summary, "Story")
    if existing:
        print(f"      Story already exists: {existing}  \"{summary}\"")
        return existing

    body = {
        "fields": {
            "project":     {"key": project_key},
            "summary":     summary,
            "description": adf_acceptance_criteria(ac_lines),
            "issuetype":   {"name": "Story"},
            "parent":      {"key": epic_key},
        }
    }
    result = jira_post("/issue", body)
    key = result["key"]
    print(f"      Created story: {key}  \"{summary}\"")
    time.sleep(0.3)
    return key


# ---------------------------------------------------------------------------
# Transition ID discovery
# ---------------------------------------------------------------------------

def get_transition_ids(project_key):
    """Return a dict of {transition_name: id} for the first issue found."""
    issues = search_issues(f'project = "{project_key}"', fields="summary")
    if not issues:
        return {}
    issue_key = issues[0]["key"]
    data = jira_get(f"/issue/{issue_key}/transitions")
    return {t["name"]: t["id"] for t in data.get("transitions", [])}


# ---------------------------------------------------------------------------
# README generation
# ---------------------------------------------------------------------------

def write_readme(project_key, epics_map, transition_ids):
    lines = [
        f"# {PROJECT_NAME}",
        "",
        PROJECT_DESC,
        "",
        "---",
        "",
        "## Board Guide",
        "",
        f"- **Project key:** `{PROJECT_KEY}`",
        f"- **Board URL:** {JIRA_BASE_URL}/jira/software/projects/{PROJECT_KEY}/boards",
        "- **Workflow:** To Do → In Progress → Done",
        "",
        "---",
        "",
        "## Epics",
        "",
    ]

    for epic_summary, epic_key in epics_map.items():
        lines.append(f"| [{epic_key}]({JIRA_BASE_URL}/browse/{epic_key}) | {epic_summary} |")

    lines += [
        "",
        "---",
        "",
        "## JQL Examples",
        "",
        "```jql",
        f"# All open stories",
        f'project = {PROJECT_KEY} AND issuetype = Story AND statusCategory != Done ORDER BY created ASC',
        "",
        f"# Stories in a specific epic (replace AVBA-1 with actual epic key)",
        f'"Epic Link" = AVBA-1 ORDER BY created ASC',
        "",
        f"# Everything not yet started",
        f'project = {PROJECT_KEY} AND status = "To Do" ORDER BY created ASC',
        "",
        f"# In-progress items",
        f'project = {PROJECT_KEY} AND status = "In Progress" ORDER BY updated DESC',
        "```",
        "",
        "---",
        "",
        "## Transition IDs",
        "",
        "Use these with `POST /rest/api/3/issue/{issueKey}/transitions`.",
        "",
        "| Transition Name | ID |",
        "|---|---|",
    ]

    if transition_ids:
        for name, tid in sorted(transition_ids.items()):
            lines.append(f"| {name} | {tid} |")
    else:
        lines.append("| *(run script with --transitions flag after first issue is created)* | - |")

    lines += [
        "",
        "---",
        "",
        "## Setup Script",
        "",
        "```bash",
        "pip install requests python-dotenv",
        "cp .env.example .env   # fill in credentials",
        "python setup_jira.py",
        "```",
        "",
        "The script is idempotent — safe to re-run if it fails partway through.",
        "",
        "---",
        "",
        "## Security Note",
        "",
        "> Never commit your `.env` file. It is listed in `.gitignore`.",
        "> Rotate your Jira API token at https://id.atlassian.net/manage-profile/security/api-tokens",
        "> if it has ever appeared in plain text in a chat, email, or terminal log.",
    ]

    readme_path = os.path.join(os.path.dirname(__file__), "README.md")
    with open(readme_path, "w") as f:
        f.write("\n".join(lines) + "\n")
    print(f"\n  README written to {readme_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=== AccessVBA → Azure SQL Migration — Jira Setup ===\n")

    # 1. Project
    print("[1/3] Project")
    project_key, _ = get_or_create_project()

    # 2. Epics + Stories
    print("\n[2/3] Epics and Stories")
    epics_map = {}  # {epic_summary: epic_key}

    for block in BOARD:
        print(f"\n  Epic: {block['epic_summary']}")
        epic_key = get_or_create_epic(
            project_key,
            block["epic_summary"],
            block["epic_description"],
        )
        epics_map[block["epic_summary"]] = epic_key

        for story_summary, ac_lines in block["stories"]:
            get_or_create_story(project_key, epic_key, story_summary, ac_lines)

    # 3. README
    print("\n[3/3] Transition IDs + README")
    transition_ids = get_transition_ids(project_key)
    if transition_ids:
        print("  Transitions found:")
        for name, tid in transition_ids.items():
            print(f"    {tid:>6}  {name}")
    else:
        print("  No issues exist yet to query transitions from.")

    write_readme(project_key, epics_map, transition_ids)

    print("\nDone.")
    print(f"Board: {JIRA_BASE_URL}/jira/software/projects/{project_key}/boards")


if __name__ == "__main__":
    main()

# AccessVBA → Azure SQL Migration — Jira Board Guide

Migrate the Q-1019 Access database backend from a shared-drive .accdb to Azure SQL
(AAD auth), and move frontend distribution to SharePoint.

---

## Quick Start

```bash
pip install requests python-dotenv
cp .env.example .env        # fill in your credentials (see Security Note below)
python setup_jira.py
```

The script creates the project, 4 epics, and 19 user stories. It is **idempotent** —
safe to re-run if it fails partway through.

---

## Project

| Field | Value |
|---|---|
| Project name | AccessVBA to Azure SQL Migration |
| Project key | `AVBA` |
| Type | Software / Kanban |
| Board URL | `https://blkbrd77.atlassian.net/jira/software/projects/AVBA/boards` |

---

## Epics & Stories

### Epic 1 — Azure SQL Setup & Schema
| Story | Acceptance Criteria (summary) |
|---|---|
| Provision Azure SQL Database instance | Instance created, firewall rules set, connection confirmed from dev machine |
| Migrate Access schema using SSMA | All tables, indexes, and FKs present; zero SSMA critical errors |
| Configure Azure AD authentication | AAD group created, least-privilege login, 2+ test users verified |
| Validate network access (office + remote) | Both environments reach Azure SQL; latency < 500 ms remote |

### Epic 2 — Data Migration
| Story | Acceptance Criteria (summary) |
|---|---|
| Run initial data load via SSMA | Row counts match Access BE for every table |
| Validate order numbering integrity | No gaps/duplicates in SalesOrders; OrderSeq values correct |
| Validate reference and lookup tables | All lookup tables present, spot-check 10 records each |
| Cutover delta sync and lock old BE | Delta synced, old .accdb renamed to .bak, Azure SQL is sole source |

### Epic 3 — FE Re-link & VBA Updates
| Story | Acceptance Criteria (summary) |
|---|---|
| Update basRelinkTables for ODBC | RelinkAllTables() re-links all tables via DSN-less ODBC string |
| Update basRemoteAccess for ODBC | TestBackendConnection() tests ODBC, not file path |
| Update tblConfig for Azure SQL keys | New keys (AzureSQLServer, AzureSQLDatabase, AzureSQLAuthMode); no credentials stored |
| Update basSeqAllocator for SQL Server | No duplicate order numbers under 5-user concurrent load |
| Remediate Access-specific SQL | All saved queries audited; incompatible ones rewritten as passthrough |

### Epic 4 — FE Distribution via SharePoint
| Story | Acceptance Criteria (summary) |
|---|---|
| Set up SharePoint library for FE template | Library created, versioning on, correct permissions |
| Update basVersion for SharePoint update detection | Startup check prompts user if newer version exists; fails gracefully |
| Write user installation and update guide | Covers first install, update flow, ODBC Driver 18 prerequisite |
| Pilot rollout (2-3 users) | One office + one remote user confirm end-to-end order creation |
| Full rollout and decommission shared-drive FEs | All users on SharePoint FE; old copies removed |

---

## JQL Examples

```jql
-- All open stories
project = AVBA AND issuetype = Story AND statusCategory != Done ORDER BY created ASC

-- Stories under a specific epic (replace AVBA-1 with the actual key)
"Epic Link" = AVBA-1 ORDER BY created ASC

-- Everything not yet started
project = AVBA AND status = "To Do" ORDER BY created ASC

-- In-progress items
project = AVBA AND status = "In Progress" ORDER BY updated DESC

-- All epics
project = AVBA AND issuetype = Epic ORDER BY created ASC

-- Stories with no assignee
project = AVBA AND issuetype = Story AND assignee is EMPTY AND statusCategory != Done
```

---

## Transition IDs

Query live transition IDs for any issue with:

```bash
curl -u "you@example.com:YOUR_API_TOKEN" \
  -H "Accept: application/json" \
  "https://blkbrd77.atlassian.net/rest/api/3/issue/AVBA-1/transitions"
```

Typical Kanban transitions (IDs vary per project — verify after creation):

| Transition | Typical ID |
|---|---|
| Start Progress | 21 |
| Done | 31 |
| Stop Progress | 11 |

To move an issue programmatically:

```bash
curl -u "you@example.com:YOUR_API_TOKEN" \
  -X POST \
  -H "Content-Type: application/json" \
  -d '{"transition": {"id": "21"}}' \
  "https://blkbrd77.atlassian.net/rest/api/3/issue/AVBA-5/transitions"
```

---

## Security Note

> **Never commit your `.env` file.** It is listed in `.gitignore`.
>
> If your API token has appeared in plain text anywhere (chat, email, terminal log),
> rotate it immediately at:
> https://id.atlassian.net/manage-profile/security/api-tokens

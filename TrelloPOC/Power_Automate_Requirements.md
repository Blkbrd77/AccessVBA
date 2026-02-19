# Power Automate Flow — Requirements & Field Mapping

Fill in each blank below. Once complete, this document will drive the full flow build.

---

## 1. Placker API — Confirmed Field Names

*(From your Postman results — open the Console tab for each request)*

### Card Detail (`GET /card/{card}`)

| Required Data | Exact API Field Name | Example Value Seen |
|---|---|---|
| Card title / name | ________________________ | |
| Close / completion date | ________________________ | |
| Card description | ________________________ | |
| Card ID | ________________________ | |

### Comments (`GET /card/{card}/comment`)

| Required Data | Exact API Field Name | Example Value Seen |
|---|---|---|
| Comment text/body | ________________________ | |
| Comment author | ________________________ | |
| Comment date | ________________________ | |

### Checklists (`GET /card/{card}/checklist`)

| Required Data | Exact API Field Name | Example Value Seen |
|---|---|---|
| Checklist name | ________________________ | |
| Item name/text | ________________________ | |
| Item completion state | ________________________ | e.g., `complete` / `incomplete` |

### Webhook Payload (`GET /webhook/{board}/example`)

| Required Data | Exact Field Name in Payload | Notes |
|---|---|---|
| Card ID (used in follow-up calls) | ________________________ | |
| Event type (to detect "moved to Done") | ________________________ | |
| List name in payload | ________________________ | |

---

## 2. SharePoint — Destination

| Question | Your Answer |
|---|---|
| SharePoint site URL | ________________________ |
| List or Document Library name | ________________________ |
| How is it organized? (by customer folder / by row / other) | ________________________ |

### SharePoint Column Names to Write To

*(List every column in SharePoint that the flow needs to populate)*

| Column Name in SharePoint | Data Source | Notes |
|---|---|---|
| ________________________ | Card title | |
| ________________________ | Close date | |
| ________________________ | Comments | Plain text or multi-line? |
| ________________________ | Checklist completion % | Calculated or raw? |
| ________________________ | ________________________ | |
| ________________________ | ________________________ | |

---

## 3. Trigger — How the Flow Starts

| Question | Your Answer |
|---|---|
| Trigger type | Placker webhook  /  Scheduled poll  /  Manual |
| If webhook: what event fires it? (e.g., card moved to list) | ________________________ |
| If scheduled: how often? | ________________________ |
| Should the flow filter to a specific board? | Yes — Board ID: ____________  /  No |
| Should the flow filter to a specific list name? | Yes — List name: __________  /  No |

---

## 4. SharePoint — Matching Logic

*(How does the flow know which SharePoint row/folder to update?)*

| Question | Your Answer |
|---|---|
| What field on the card matches a SharePoint record? | ________________________ |
| Is it a lookup by card name, card ID, a custom field? | ________________________ |
| What happens if no match is found? | Create new row  /  Skip  /  Flag error |
| What happens if multiple matches are found? | Update first  /  Flag error  /  Other |

---

## 5. Comments — Handling

| Question | Your Answer |
|---|---|
| Write all comments or only the most recent? | All  /  Most recent: ______ |
| Format: concatenated into one field, or one row per comment? | One field  /  One row per comment |
| Include author and date with each comment? | Yes  /  No |

---

## 6. Checklist — Handling

| Question | Your Answer |
|---|---|
| Write checklist as a completion % (e.g., 4/5 = 80%)? | Yes  /  No |
| Write full item list with status? | Yes  /  No |
| Which checklists — all on the card, or a specific one? | All  /  Specific name: __________ |

---

## 7. Error Handling & Notifications

| Question | Your Answer |
|---|---|
| Who should be notified if the flow fails? | ________________________ |
| Notification method | Email  /  Teams message  /  None |
| Should failed runs be logged somewhere? | Yes — where: ______________  /  No |

---

## 8. Placker Credentials in Power Automate

| Question | Your Answer |
|---|---|
| How will the API key be stored? | Custom connector  /  Environment variable  /  Hardcoded (not recommended) |
| Is there a shared service account for the flow? | Yes — account: ______________  /  No, runs under my account |

---

## Notes / Anything Else

```
[Add anything not covered above — edge cases, business rules, existing flows to connect to, etc.]




```

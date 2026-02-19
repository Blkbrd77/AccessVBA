# Placker API Setup & Postman Testing Guide

**Task:** Validate the Placker API to confirm it returns comments, checklist completion, and close date for Trello cards moved to "Done."

**Reference Docs:** https://placker.com/docs/api/index.html

---

## Step 1: Generate Your Placker API Key

1. Log in to Placker (https://placker.com)
2. Navigate to your **Profile > Settings > API** (or look for an "API Keys" section)
3. Click **Generate API Key**
4. Copy the key — treat it like a password, do not share it

---

## Step 2: Import the Postman Collection

1. Open Postman
2. Click **Import** (top left)
3. Select the file: `Placker_API_Tests.postman_collection.json`
4. The collection "Placker API - POC Validation" will appear in your sidebar

---

## Step 3: Set Collection Variables

1. Click the collection name in the sidebar
2. Go to the **Variables** tab
3. Fill in:

| Variable | Value |
|---|---|
| `baseUrl` | Verify against the API docs — likely `https://placker.com/api/v1` |
| `apiKey` | Your key from Step 1 |
| `boardId` | Leave blank — populated after running Request 2 |
| `listId` | Leave blank — populated after running Request 3 |
| `cardId` | Leave blank — auto-populated by Request 4 tests |

4. Click **Save**

---

## Step 4: Run Requests in Order

### Request 1 — Auth Check
**Purpose:** Confirm the API key works and the endpoint is reachable.

- **Pass:** 200 OK response with your user account data
- **Fail (401):** API key is wrong or not formatted correctly — check the Authorization header format; it may be `X-API-Key: {{apiKey}}` instead of `Bearer`
- **Fail (403):** Your Placker plan may not include API access

### Request 2 — List Boards
**Purpose:** Find the board ID for your Trello board.

- Run the request
- In the Postman console (View > Console), find your target board
- Copy its `id` value into the `boardId` collection variable

### Request 3 — Get Lists on Board
**Purpose:** Find the "Done" list ID.

- Run the request with `boardId` set
- Find the list named "Done" (or your equivalent completion list)
- Copy its `id` into the `listId` variable

### Request 4 — Get Cards in Done List
**Purpose:** Retrieve cards from Done and inspect data shape.

- Run the request
- The test script logs the full first card to the Postman console
- The `cardId` variable is auto-set to the first card's ID
- **Key:** Note which top-level fields are present for comments, checklists, and dates

### Request 5 — Get Single Card Detail
**Purpose:** Primary validation of the 3 required data fields.

- Check the **Test Results** tab after running
- Three tests will run:
  1. **REQUIREMENT CHECK: Comments field present**
  2. **REQUIREMENT CHECK: Checklist completion field present**
  3. **REQUIREMENT CHECK: Close/completion date field present**
- The full card JSON is logged to console for inspection

### Request 6 — Get Card Comments (Explicit)
**Purpose:** Determine if comments have a dedicated sub-endpoint or are embedded.

- If 200: document the `comments` sub-resource path for Power Automate
- If 404: use the embedded comments field found in Request 5

---

## Step 5: Document Findings

After running the tests, complete this table for the Power Automate mapping:

| Required Field | API Field Name | Endpoint | Notes |
|---|---|---|---|
| Comments | `???` | `/cards/{id}` or `/cards/{id}/comments` | |
| Checklist completion | `???` | `/cards/{id}` | Note if items have `state: complete/incomplete` |
| Close date | `???` | `/cards/{id}` | Note exact field name (e.g., `closedDate`, `actualEndDate`) |

---

## Troubleshooting

### Wrong base URL
The API docs are at https://placker.com/docs/api/index.html — if requests fail with 404, check the actual base URL. It may be `https://app.placker.com/api/v1` or similar.

### Auth header format
If Bearer token auth fails, try:
- `X-API-Key: {{apiKey}}`
- `Api-Key: {{apiKey}}`
- `apiKey` as a query parameter: `?apiKey={{apiKey}}`

Check the Placker docs auth section for the correct format.

### Rate limiting
If you receive 429 responses, wait and retry. Add delays between requests in Postman if running as a collection runner.

---

## Next Steps (After API Validation)

Once the 3 fields are confirmed:

1. **Document the field names** in the table above
2. **Design the Power Automate flow:**
   - Trigger: HTTP webhook from Placker when a card moves to "Done"
   - Action 1: Call Placker API to get full card detail
   - Action 2: Extract comments, checklist completion %, close date
   - Action 3: Write to the matching customer folder in SharePoint
3. **Test the webhook** using Request 7 (to be added) or via Placker webhook settings

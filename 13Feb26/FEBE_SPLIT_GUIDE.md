# Q1019 Database — Frontend/Backend Split Guide

**Date:** February 2026
**Database:** `12Feb26FE_BESplit.accdb`
**Target:** `Q1019_BE.accdb` (backend) + `Q1019_FE.accdb` (frontend)

---

## Overview

A split database means:

- **Backend (BE):** One file on the server containing all data tables. Nobody opens this
  directly — it is always accessed through linked tables.
- **Frontend (FE):** One file per user, stored locally on each user's machine. Contains
  all forms, reports, queries, and VBA modules. Has no local data tables — only linked
  tables pointing to the BE.

```
Each user's local machine          Shared server
+-----------------------+          +------------------------+
|  Q1019_FE.accdb       | <--LAN-> |  Q1019_BE.accdb        |
|  Forms, reports       |   or VPN |  All data tables       |
|  Modules, queries     |          |  OrderSeq, SalesOrders  |
|  Linked tables -------+----------+- tblConfig, etc.       |
+-----------------------+          +------------------------+
```

---

## Before You Begin

- [ ] Close the database on all machines -- you must be the only user open
- [ ] Make a full backup: copy `12Feb26FE_BESplit.accdb` to a safe location with today's
      date in the filename
- [ ] Confirm you have write access to the server folder where the BE will live
- [ ] Set aside 30-45 minutes of uninterrupted time
- [ ] Know the exact UNC path where the BE will live, e.g.:
      `\\hmc-dc01\Users\JSAMPLES\Q-1019_Database\`

---

## Step 1: Run the Database Splitter Wizard

This is the built-in Access tool. It does the heavy lifting automatically.

1. Open `12Feb26FE_BESplit.accdb`
2. Click the **Database Tools** tab on the ribbon
3. Click **Move Data** → **Access Database**
4. The wizard says *"This wizard moves tables from your current database to a new
   back-end database..."* — click **Split Database**
5. A **Save As** dialog opens. Navigate to your server folder:
   `\\hmc-dc01\Users\JSAMPLES\Q-1019_Database\`
6. Name the file: `Q1019_BE.accdb`
7. Click **Split**
8. Wait 30-60 seconds. When done you will see: *"Database successfully split."*
9. Click **OK**

**What the wizard just did:**
- Created `Q1019_BE.accdb` on the server with all your data tables
- In your current file, deleted all local tables and replaced them with **linked tables**
  pointing to the BE
- Your current file is now effectively the FE

---

## Step 2: Rename Your File

1. **Close** the database (File → Close)
2. In File Explorer, navigate to your working copy
3. Rename `12Feb26FE_BESplit.accdb` → `Q1019_FE.accdb`
4. Do not move it yet — keep it local while you test

---

## Step 3: Verify the Split Worked

1. Open `Q1019_FE.accdb`
2. In the Navigation Pane, expand **Tables**
3. Every table should show a **small arrow icon** (→) — that means it is a linked table,
   not a local one
4. Double-click a table (e.g., SalesOrders) — it should open and show your data
5. Open frmAbout — the **Backend Path** field should show the path to `Q1019_BE.accdb`

---

## Step 4: Update tblConfig With the Correct Paths

Open `tblConfig` and set or add these rows:

| ConfigKey        | ConfigValue                                                           |
|------------------|-----------------------------------------------------------------------|
| `BackendPath`    | `\\hmc-dc01\Users\JSAMPLES\Q-1019_Database\Q1019_BE.accdb`           |
| `FE_TemplatePath`| `\\hmc-dc01\Users\JSAMPLES\Q-1019_Database\Q1019_FE_TEMPLATE.accdb`  |

1. Open `tblConfig` directly from the Navigation Pane
2. Find (or add) the `BackendPath` row
3. Set `ConfigValue` to the full UNC path of `Q1019_BE.accdb`
4. Save and close

---

## Step 5: Replace basRemoteAccess

The simplified `basRemoteAccess` module (see `basRemoteAccess.bas` in this folder) must
be imported to replace the old one **before** the next step.

1. In Access: Alt+F11 to open the VBA Editor
2. In the Project Explorer, right-click `basRemoteAccess` → **Remove basRemoteAccess**
   (say No when asked to export)
3. File → **Import File** → navigate to `13Feb26/basRemoteAccess.bas` → Open

**Also required in Form_frmOrderList:**
Open `frmOrderList` in Design View → View Code and delete this block from `Form_Load`:

```vba
' DELETE these 6 lines:
If IsRemoteConnection() Then
    MsgBox "Remote connection detected (VPN/WAN latency)." & vbCrLf & _
           "For best performance: keep this app as your only Access instance, " & _
           "avoid leaving forms open unnecessarily, and save frequently.", _
           vbInformation, "Remote Access"
End If
```

Everything else in `Form_frmOrderList` (the `EnsureBackendConnectionOrExit` function and
the `EH:` block that calls `HandleNetworkError`) is unchanged and works with the new
module.

---

## Step 6: Test the Connection Check

In the VBA Immediate Window (Alt+F11, then Ctrl+G):

```vba
?TestBackendConnection()
```

Should return `True`. Then confirm a log entry was written:

```vba
?DLookup("Details","tblConcurrencyLog","EventType='BackendConnection'")
```

Should show: `PASS: Connected in 42 ms. Path=\\hmc-dc01\...`

---

## Step 7: Test All Core Functionality

Before distributing to users:

- [ ] Open frmOrderList — loads without error
- [ ] Create a test batch (Preview mode — do not commit)
- [ ] Confirm sequence numbers generate correctly
- [ ] Open frmAbout — all fields populated, backend path correct
- [ ] Check tblConcurrencyLog — entries are being written

---

## Step 8: Create the Template Copy for Users

Users each need their own local FE. Keep a "template" on the server for distribution.

1. Copy `Q1019_FE.accdb` to the server:
   `\\hmc-dc01\Users\JSAMPLES\Q-1019_Database\Q1019_FE_TEMPLATE.accdb`
2. This template stays on the server — it is not what users open to work
3. When a user needs a fresh copy, they copy from here

---

## Step 9: Distribute the Frontend to Each User

### Local Network Users (4 users)

1. On the user's machine, create: `C:\Users\[Username]\Documents\Q1019\`
2. Copy `Q1019_FE_TEMPLATE.accdb` from the server to that folder
3. Rename the copy to `Q1019_FE.accdb`
4. Create a Desktop shortcut pointing to that file
5. Have the user open from the shortcut and test

### Remote VPN Users (2 users)

Same steps as above, plus communicate this:

> **Important:** You must be connected to VPN **before** opening the database.
> The database connects to `\\hmc-dc01\Users\JSAMPLES\Q-1019_Database\Q1019_BE.accdb`
> — this path is only reachable on VPN. If you open without VPN you will see a
> "Cannot connect to the database" message. Simply close Access, connect VPN, and reopen.

---

## Step 10: Verify Multi-User Access

With at least 2 users connected at the same time:

1. Both open the database
2. One user opens a form; the other opens the same form — no errors
3. One user creates a test order while the other is active
4. After refresh, both users see the new order
5. Confirm no lock errors

---

## Step 11: Lock Down the Backend

1. On the server, confirm regular users have **Read/Write** but **not Full Control**
   on the `Q1019_BE.accdb` file
2. Users should not be able to delete the BE file
3. Only the admin account needs Full Control

---

## Ongoing Maintenance

### Updating the Frontend (Bug Fix or New Feature)

1. Make and test the change in your development copy
2. Bump the version in `tblAppVersion`
3. Overwrite `Q1019_FE_TEMPLATE.accdb` on the server with the updated FE
4. Notify users to close, delete their local FE, copy the new template, and reopen

### Updating the Backend (Schema Change)

- **New table added:** After the BE change, go into each FE and run Linked Table Manager
  (External Data → Linked Table Manager → Select All → Relink)
- **Data or index change only:** No FE action needed

### Daily Backup

The BE is the only file with your data. Back it up daily:

```vba
Public Sub BackupBackend()
    Dim src  As String
    Dim dest As String
    src  = GetConfig("BackendPath")
    dest = GetConfig("BackupPath") & "Q1019_BE_" & Format(Now(), "yyyymmdd_hhnnss") & ".accdb"
    FileCopy src, dest
    MsgBox "Backup complete: " & dest, vbInformation
End Sub
```

### Monthly Compact of the Backend

1. Confirm no users are connected (no `Q1019_BE.ldb` file in the server folder)
2. Open Access (not the database — just Access)
3. File → Open → browse to `Q1019_BE.accdb`
4. Database Tools → Compact & Repair Database
5. Close

---

## Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| "Could not find file" on FE open | Linked tables can't reach the BE | Check network/VPN; verify UNC path in tblConfig |
| Tables show `#Name?` or `#Deleted` | Broken linked tables | External Data → Linked Table Manager → Relink all |
| "Database is locked by another user" | `.ldb` file left by crashed session | Confirm no users connected; delete the `.ldb` |
| Remote user: "Cannot connect" at startup | VPN not connected | Close Access, connect VPN, verify `\\hmc-dc01\` in Explorer, reopen |
| `TestBackendConnection()` returns False but file exists | BackendPath typo in tblConfig | Open tblConfig, verify the ConfigValue exactly matches the real file path |

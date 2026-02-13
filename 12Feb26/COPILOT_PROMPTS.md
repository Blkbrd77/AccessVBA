# Copilot Professional Prompts - Ready to Copy & Paste
**Q1019 Database Multi-User Implementation**
**Date: February 12, 2026**

---

## How to Use This Document

1. Find the prompt you need
2. Copy the entire prompt (between the ``` markers)
3. Paste into Copilot Professional
4. Review and implement the generated code
5. Test thoroughly before moving to next prompt

**⚠️ Important:** Do prompts in order! Some depend on previous work.

---

## Table of Contents

### Phase 1: CRITICAL (Must Do First)
- [1.1 Pessimistic Locking](#11-pessimistic-locking-for-reserveseq)
- [1.2 Concurrency Logging](#12-concurrency-logging-module)

### Phase 2: HIGH PRIORITY
- [2.1 Configuration System](#21-configuration-system)
- [2.2 Version Tracking](#22-version-tracking-system)
- [2.3 Remote Access Handling](#23-remote-access-module)

### Phase 3: NICE TO HAVE
- [3.1 Multi-User Test Suite](#31-multi-user-test-suite)
- [3.2 Admin Dashboard](#32-admin-dashboard-optional)

---

# Phase 1: CRITICAL - Do These First

## 1.1 Pessimistic Locking for ReserveSeq

**Purpose:** Fix race condition that causes duplicate order numbers

**Copy this prompt:**

```
I have an Access VBA module called basSeqAllocator with a ReserveSeq function
that has a race condition for multi-user environments.

The function currently opens a recordset without locking, which means multiple
users can read the same NextSeq value simultaneously and create duplicate order
numbers.

UPDATE THE RESERVESEQ FUNCTION TO USE PESSIMISTIC LOCKING:

REQUIREMENTS:
1. Add Windows Sleep API declaration (handle both 32-bit and 64-bit Office):
   - #If VBA7 Then / #Else for compatibility
   - Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

2. Change line 1987 from:
   Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)

   To use pessimistic locking:
   Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbPessimistic)

3. Implement retry logic:
   - Maximum 5 retry attempts
   - Wait 500ms between retries using Sleep API
   - Handle error 3260 (record locked by another user)
   - After 5 failed attempts, raise clear error to user

4. Preserve existing functionality:
   - IsPreview parameter still works (doesn't increment if True)
   - Function signature unchanged
   - Returns Long (sequence number)

5. Add documentation:
   - Comment explaining the locking mechanism
   - Comment explaining retry logic
   - Comment for each major code block

EXISTING FUNCTION SIGNATURE (DO NOT CHANGE):
Public Function ReserveSeq( _
    ByVal Scope As String, _
    ByVal BaseToken As String, _
    ByVal QualifierCode As String, _
    ByVal SystemLetter As String, _
    Optional ByVal IsPreview As Boolean = False _
) As Long

ERROR HANDLING:
- Use On Error GoTo ErrorHandler pattern
- Check for error 3260 specifically (locked record)
- For error 3260: retry up to 5 times with 500ms sleep
- For other errors: raise immediately with clear message
- After max retries: raise custom error with message:
  "Could not reserve sequence after 5 attempts. Another user is currently
   reserving sequences. Please try again in a moment."

PROVIDE:
- Complete updated basSeqAllocator module code
- Sleep API declaration at module level
- Updated ReserveSeq function with pessimistic locking and retry logic
```

**After implementation:**
- Test: `?ReserveSeq("TEST", "TST", "", "T", False)` in Immediate Window
- Should return sequence number without errors

---

## 1.2 Concurrency Logging Module

**Purpose:** Track multi-user issues for monitoring and troubleshooting

**Copy this prompt:**

```
Create a VBA module called basAuditLog for logging concurrency events in a
multi-user Access database.

CREATE TABLE:
Table name: tblConcurrencyLog

Fields:
- LogID: AutoNumber (Primary Key)
- LogTimestamp: Date/Time (indexed, default = Now())
- EventType: Text(50) (values like "SeqReservationRetry", "LockTimeout")
- Details: Text(255) (event description)
- UserName: Text(100) (Windows username)
- ComputerName: Text(100) (computer name)

CREATE MODULE: basAuditLog

Function signature:
Public Sub LogConcurrencyEvent( _
    ByVal EventType As String, _
    ByVal Details As String, _
    Optional ByVal UserName As String = "" _
)

Function behavior:
1. If UserName parameter is empty, get from Environ("USERNAME")
2. Get ComputerName from Environ("COMPUTERNAME")
3. Create table if doesn't exist (use CurrentDb.TableDefs)
4. Open recordset and add new log entry
5. Use On Error Resume Next (logging failure shouldn't crash app)
6. Close recordset properly

CREATE QUERY: qryConcurrencyLog_Last24Hours

SQL:
SELECT
    LogID,
    Format(LogTimestamp, "yyyy-mm-dd hh:nn:ss") AS EventTime,
    EventType,
    Details,
    UserName,
    ComputerName
FROM tblConcurrencyLog
WHERE LogTimestamp >= DateAdd("h", -24, Now())
ORDER BY LogTimestamp DESC;

INTEGRATION:
Update basSeqAllocator.ReserveSeq to call LogConcurrencyEvent when retries occur.

In the retry loop, after error 3260 is caught, before Sleep:
    LogConcurrencyEvent "SeqReservationRetry", _
        "Attempt " & retryCount & " of 5 for Scope=" & Scope & _
        ", BaseToken=" & BaseToken & ", Qualifier=" & QualifierCode

PROVIDE:
- Complete basAuditLog module code with table creation
- LogConcurrencyEvent function
- SQL for query
- Updated ReserveSeq with logging integration
```

**After implementation:**
- Test: `LogConcurrencyEvent "TEST", "Testing logging system"`
- Open query qryConcurrencyLog_Last24Hours - should see test entry

---

# Phase 2: HIGH PRIORITY

## 2.1 Configuration System

**Purpose:** Centralized settings management for database paths and parameters

**Copy this prompt:**

```
Create a configuration management system for a multi-user Access database.

CREATE TABLE: tblConfig

Fields:
- ConfigKey: Text(50) - Primary Key
- ConfigValue: Text(255)
- Description: Text(255)
- LastUpdated: Date/Time
- UpdatedBy: Text(100)

POPULATE WITH INITIAL VALUES:
INSERT these configuration values:

ConfigKey | ConfigValue | Description
----------|-------------|-------------
BackendPath | \\server\share\Q1019\Backend\Q1019_BE.accdb | Path to backend database
FE_TemplatePath | \\server\share\Q1019\FrontEnd\Q1019_FE_TEMPLATE.accdb | Frontend template for updates
BackupPath | \\server\share\Q1019\Backups\ | Backup storage location
MinRequiredVersion | 1.0.0 | Minimum frontend version allowed
MaxConcurrentUsers | 6 | Maximum simultaneous users
EnableRemoteAccess | Yes | Allow VPN access
LogRetentionDays | 30 | Days to keep concurrency logs
SequenceRetryAttempts | 5 | Max retries for sequence locks
SequenceRetryDelayMs | 500 | Milliseconds between retries
AdminEmail | admin@company.com | Contact for issues

(Note: Replace network paths with actual paths for your environment)

CREATE MODULE: basConfig

Functions needed:
1. Public Function GetConfig(ByVal ConfigKey As String) As String
   - Returns ConfigValue for given key
   - Returns default value if key not found
   - Handle errors silently

2. Public Function GetConfigBoolean(ByVal ConfigKey As String) As Boolean
   - Returns True if value is "Yes", "True", or "1"
   - Returns False otherwise

3. Public Function GetConfigNumber(ByVal ConfigKey As String) As Long
   - Returns ConfigValue converted to Long
   - Returns 0 if conversion fails

4. Public Sub SetConfig(ByVal ConfigKey As String, ByVal ConfigValue As String)
   - Updates existing key or adds new key
   - Sets LastUpdated = Now()
   - Sets UpdatedBy = Environ("USERNAME")

ERROR HANDLING:
- Use On Error Resume Next for Get functions (return defaults)
- Use On Error GoTo for Set function (show error to user)

DEFAULT VALUES (if key not found):
- SequenceRetryAttempts: "5"
- SequenceRetryDelayMs: "500"
- LogRetentionDays: "30"
- MaxConcurrentUsers: "6"
- Others: empty string ""

PROVIDE:
- SQL to create table
- SQL to insert initial values (as individual INSERT statements)
- Complete basConfig module code with all four functions
```

**After implementation:**
- Test: `?GetConfig("BackendPath")`
- Test: `?GetConfigNumber("SequenceRetryAttempts")`
- Test: `?GetConfigBoolean("EnableRemoteAccess")`

---

## 2.2 Version Tracking System

**Purpose:** Track frontend versions for updates and compatibility

**Copy this prompt:**

```
Create a version tracking system for Access frontend with version comparison capability.

CREATE TABLE: tblAppVersion

Fields:
- VersionNumber: Text(20) - e.g., "1.0.0" (semantic versioning)
- ReleaseDate: Date/Time - when released
- ReleaseNotes: Memo/LongText - what changed
- IsActive: Yes/No - current version flag
- MinBackendVersion: Text(20) - required backend version

INSERT INITIAL VERSION:
VersionNumber: 1.0.0
ReleaseDate: #2026-02-12#
ReleaseNotes: Initial multi-user release with pessimistic locking, concurrency
              logging, configuration system, and remote access support. This
              version prevents duplicate order numbers and supports 6 concurrent
              users including 2 remote admins over VPN.
IsActive: True
MinBackendVersion: 1.0.0

CREATE MODULE: basVersion

Functions:
1. Public Function GetCurrentVersion() As String
   - Returns VersionNumber where IsActive = True
   - Returns "0.0.0" if no active version found

2. Public Sub SetVersion(ByVal versionNumber As String, ByVal releaseNotes As String)
   - Sets all IsActive to False
   - Adds new version with IsActive = True
   - Sets ReleaseDate = Now()
   - Sets MinBackendVersion to same as versionNumber

3. Public Function CompareVersions(ByVal v1 As String, ByVal v2 As String) As Long
   - Parse semantic versioning: major.minor.patch
   - Split on "." to get three numbers
   - Compare major, then minor, then patch
   - Return -1 if v1 < v2
   - Return 0 if v1 = v2
   - Return 1 if v1 > v2
   - Example: CompareVersions("1.0.0", "1.1.0") returns -1

4. Public Function IsUpdateAvailable() As Boolean
   - Get local version from GetCurrentVersion()
   - Get template path from GetConfig("FE_TemplatePath")
   - Open template database (read-only, check for errors)
   - Get template version from tblAppVersion
   - Close template database
   - Compare versions using CompareVersions
   - Return True if template version > local version
   - Return False if same or local is newer
   - Handle errors (template not accessible = return False)

5. Public Function GetLatestVersionInfo() As String
   - Get most recent version from tblAppVersion (ORDER BY ReleaseDate DESC)
   - Return formatted string:
     "Version [number]" & vbCrLf &
     "Released: [date]" & vbCrLf & vbCrLf &
     "[release notes]"

CREATE FORM: frmAbout

Controls:
- Label: "Q1019 Order Management System" (Title, large font, bold)
- Label: lblVersion (displays "Version 1.0.0")
- Label: lblReleaseDate (displays "Released: February 12, 2026")
- Textbox: txtReleaseNotes (multi-line, read-only, scrollable, shows release notes)
- Label: "Frontend Path:"
- Label: lblFrontendPath (shows CurrentDb.Name)
- Label: "Backend Path:"
- Label: lblBackendPath (shows backend path from config)
- Label: "User:"
- Label: lblUserName (shows Environ("USERNAME"))
- Label: "Computer:"
- Label: lblComputerName (shows Environ("COMPUTERNAME"))
- Button: btnCheckUpdates (caption: "Check for Updates")
- Button: btnClose (caption: "Close")

Form Code:

Private Sub Form_Load()
    ' Populate version info
    Me.lblVersion = "Version " & GetCurrentVersion()

    ' Get release info from latest version
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 ReleaseDate, ReleaseNotes FROM tblAppVersion " & _
        "WHERE IsActive = True ORDER BY ReleaseDate DESC", _
        dbOpenSnapshot)

    If Not rs.EOF Then
        Me.lblReleaseDate = "Released: " & Format(rs!ReleaseDate, "mmmm d, yyyy")
        Me.txtReleaseNotes = Nz(rs!ReleaseNotes, "")
    End If
    rs.Close

    ' Populate path info
    Me.lblFrontendPath = CurrentDb.Name
    Me.lblBackendPath = GetConfig("BackendPath")
    Me.lblUserName = Environ("USERNAME")
    Me.lblComputerName = Environ("COMPUTERNAME")
End Sub

Private Sub btnCheckUpdates_Click()
    On Error GoTo EH

    If IsUpdateAvailable() Then
        MsgBox "A new version is available on the network!" & vbCrLf & vbCrLf & _
               "Please update your frontend copy to get the latest features and fixes.", _
               vbInformation, "Update Available"
    Else
        MsgBox "You have the latest version.", vbInformation, "Up to Date"
    End If

    Exit Sub
EH:
    MsgBox "Error checking for updates: " & Err.Description, vbExclamation
End Sub

Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

PROVIDE:
- SQL to create tblAppVersion table
- SQL to insert initial version 1.0.0
- Complete basVersion module with all five functions
- Complete form code for frmAbout (Form_Load and both button click events)
```

**After implementation:**
- Test: `?GetCurrentVersion()` - should show "1.0.0"
- Test: `?CompareVersions("1.0.0", "1.1.0")` - should show -1
- Test: Open frmAbout form - should display all info

---

## 2.3 Remote Access Module

**Purpose:** Handle VPN/remote connection issues gracefully

**Copy this prompt:**

```
Create remote access detection and error handling for VPN users accessing
an Access database over WAN.

CREATE MODULE: basRemoteAccess

Functions:

1. Public Function IsRemoteConnection() As Boolean
   - Get backend path from GetConfig("BackendPath")
   - Measure time to open backend database
   - Use Timer before and after
   - Calculate response time in milliseconds
   - If response time > 200ms, return True (likely remote/VPN)
   - If response time <= 200ms, return False (likely local network)
   - Open database read-only: DBEngine.OpenDatabase(path, False, True)
   - Close database immediately after timing
   - Handle errors (if can't connect, assume remote, return True)

2. Public Function TestBackendConnection() As Boolean
   - Try to open backend database (read-only)
   - Set 3 second timeout expectation
   - Return True if successful
   - Return False if timeout or error
   - Log result to concurrency log using LogConcurrencyEvent
   - Log "BackendConnectionTest" with "Success" or "Failed: [error]"
   - Use On Error GoTo for error handling

3. Public Function HandleNetworkError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String) As Boolean
   - Translate common network errors to user-friendly messages
   - Display message box with:
     * User-friendly error explanation
     * Suggested action
     * Retry and Cancel buttons (vbRetryCancel)
   - Return True if user clicks Retry
   - Return False if user clicks Cancel
   - Log error to concurrency log before showing message

ERROR TRANSLATIONS:
- Error 3151: "Cannot connect to database server. Please check your VPN connection and try again."
- Error 3343: "Database file format is not recognized. Please contact the database administrator."
- Error 3356: "Cannot find database on network. Please verify your network connection and VPN status."
- Error 3704: "Connection was lost during operation. Please reconnect to VPN and try again."
- Error -2147467259: "Network path not found. Please verify VPN is connected."
- Other errors: "Database error " & ErrorNumber & ": " & ErrorDescription & vbCrLf & "Please contact administrator if problem persists."

HELPER FUNCTION:

4. Public Function GetUserFriendlyError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String) As String
   - Private helper function
   - Use Select Case on ErrorNumber
   - Return appropriate user-friendly message
   - Used by HandleNetworkError

INTEGRATION EXAMPLE:

Add to startup form or main form:

Private Sub Form_Open(Cancel As Integer)
    ' Test backend connection on startup
    If Not TestBackendConnection() Then
        If Not HandleNetworkError(3356, "Backend not accessible") Then
            ' User cancelled, close database
            Cancel = True
            Application.Quit
        End If
    End If
End Sub

PROVIDE:
- Complete basRemoteAccess module with all functions
- Error translation logic in GetUserFriendlyError
- Example integration code for startup form
- Proper error handling throughout
```

**After implementation:**
- Test: `?IsRemoteConnection()` - shows True/False based on connection speed
- Test: `?TestBackendConnection()` - should show True if backend accessible
- Add startup check to main form

---

# Phase 3: NICE TO HAVE (If Time Permits)

## 3.1 Multi-User Test Suite

**Purpose:** Automated and manual tests for concurrent access

**Copy this prompt:**

```
Create a comprehensive test suite for multi-user Access database concurrency testing.

CREATE MODULE: basTestMultiUser

Test Functions:

1. Public Function TestSequenceReservation() As Boolean
   TEST PURPOSE: Verify single-user sequence reservation works correctly

   STEPS:
   - Print header with "=" separators
   - Reserve 100 sequences using scope "TEST", baseToken "TST"
   - Use Timer to measure elapsed time
   - Store each sequence number in array
   - Print progress every 10 sequences
   - After reserving all, check array for duplicates (nested loop)
   - Calculate average time per sequence
   - Print summary: total time, average time, duplicate status
   - Clean up: DELETE FROM OrderSeq WHERE Scope = 'TEST'
   - Return True if no duplicates, False if duplicates found

   OUTPUT:
   - Use Debug.Print for all output
   - Clear formatting with headers and separators
   - Show pass/fail status clearly

2. Public Function TestConcurrentAccess() As Boolean
   TEST PURPOSE: Instructions for manual concurrent testing

   BEHAVIOR:
   - Print detailed instructions for multi-computer testing
   - Instructions should say:
     * Copy frontend to 2-3 computers
     * Open database on each computer
     * Coordinate to run this function at EXACT same time
     * Each computer will reserve 50 sequences
     * After all complete, run ValidateOrderSequences() on one computer
   - Wait for user confirmation (MsgBox with OK button)
   - Print computer name, username, start timestamp
   - Reserve 50 sequences with scope "CONCURRENT_TEST", baseToken "CT"
   - Print each reserved sequence number
   - Print completion timestamp
   - Remind user to run ValidateOrderSequences()
   - Return True (manual validation required)

3. Public Function ValidateOrderSequences() As Boolean
   TEST PURPOSE: Check database for duplicate sequences

   CHECKS:
   a) Check SalesOrders table for duplicate OrderNumbers:
      SELECT OrderNumber, COUNT(*) AS cnt
      FROM SalesOrders
      GROUP BY OrderNumber
      HAVING COUNT(*) > 1

   b) Check CONCURRENT_TEST scope for expected total:
      SELECT NextSeq FROM OrderSeq WHERE Scope = 'CONCURRENT_TEST'
      NextSeq - 1 = total sequences reserved
      If 3 computers × 50 each = should be 150

   c) Check TEST scope if exists:
      Same logic as CONCURRENT_TEST

   OUTPUT:
   - Print results of each check
   - Show any duplicates found with details
   - Show total sequences reserved vs expected
   - Print PASS or FAIL summary

   CLEANUP:
   - Ask user if they want to clean up test data (MsgBox Yes/No)
   - If Yes: DELETE FROM OrderSeq WHERE Scope IN ('TEST', 'CONCURRENT_TEST')

   Return True if no duplicates, False if any found

4. Public Function TestRemoteVPN() As Boolean
   TEST PURPOSE: Test performance over VPN connection

   STEPS:
   - Detect if remote using IsRemoteConnection()
   - Print connection status (REMOTE or LOCAL)
   - Test backend connection and measure response time
   - Print response time and status
   - Warn if response time > 500ms (slow)
   - Reserve 10 test sequences and measure total time
   - Calculate average time per sequence
   - Print performance assessment:
     * < 500ms avg: GOOD
     * 500-1000ms: CAUTION (VPN latency)
     * > 1000ms: WARNING (very slow, contact IT)
   - Clean up test data
   - Return True if connection works (regardless of speed)

5. Public Sub RunAllTests()
   TEST PURPOSE: Run all automated tests in sequence

   BEHAVIOR:
   - Print formatted header with date/time
   - Run TestSequenceReservation()
   - Run TestRemoteVPN()
   - Print note about manual TestConcurrentAccess()
   - Print summary: all automated tests passed/failed
   - Print footer

FORMATTING:
- Use Debug.Print throughout
- Use "=========================================" for section separators
- Use "  " (2 spaces) for indented details
- Use "✓" or "SUCCESS:" for pass, "✗" or "ERROR:" for fail
- Include timestamps where relevant
- Make output readable and professional

ERROR HANDLING:
- Each test function: On Error GoTo EH
- Error handler prints error info and returns False
- Tests don't crash, they report failure gracefully

PROVIDE:
- Complete basTestMultiUser module
- All five functions with proper error handling
- Clean, formatted Debug.Print output
- Proper cleanup of test data
```

**After implementation:**
- Test: `RunAllTests` in Immediate Window
- Review output in Immediate Window
- For concurrent test, coordinate with multiple computers

---

## 3.2 Admin Dashboard (Optional)

**Purpose:** Monitoring and maintenance tools for administrators

**Copy this prompt:**

```
Create an administrator dashboard form for database monitoring and maintenance.

CREATE FORM: frmAdminDashboard

SECTIONS:

1. SYSTEM HEALTH (Top Section)
   Display:
   - Current active users (read .ldb file)
   - Backend file size (in MB)
   - Last backup date (from tblBackupLog or check backup folder)
   - Current database version

2. RECENT ACTIVITY (Middle Section)
   Last 7 days:
   - Total orders created
   - Orders per day (simple list or chart)
   - Most active users (top 3)
   - Most common qualifiers used

3. CONCURRENCY METRICS (Middle Section)
   - Total sequence reservation retries (last 24 hours)
   - Average retry count
   - Any lock timeouts
   - Network errors (last 24 hours)

4. QUICK ACTIONS (Bottom Section)
   Buttons:
   - Compact Backend
   - Create Backup
   - Refresh Table Links
   - View Concurrency Log (opens query)
   - Refresh Dashboard

QUERIES NEEDED:

Create query qryAdminDashboard_OrdersByDay:
SELECT
    Format(DateCreated, "yyyy-mm-dd") AS OrderDate,
    COUNT(*) AS OrderCount
FROM SalesOrders
WHERE DateCreated >= DateAdd("d", -7, Date())
GROUP BY Format(DateCreated, "yyyy-mm-dd")
ORDER BY Format(DateCreated, "yyyy-mm-dd");

Create query qryAdminDashboard_TopUsers:
SELECT TOP 3
    CreatedBy,
    COUNT(*) AS OrderCount
FROM SalesOrders
WHERE DateCreated >= DateAdd("d", -7, Date())
GROUP BY CreatedBy
ORDER BY COUNT(*) DESC;

Create query qryAdminDashboard_ConcurrencyStats:
SELECT
    COUNT(*) AS TotalEvents,
    SUM(IIF(EventType = 'SeqReservationRetry', 1, 0)) AS RetryCount,
    SUM(IIF(EventType = 'LockTimeout', 1, 0)) AS TimeoutCount,
    SUM(IIF(EventType LIKE '%Error%', 1, 0)) AS ErrorCount
FROM tblConcurrencyLog
WHERE LogTimestamp >= DateAdd("h", -24, Now());

FORM CODE:

Private Sub Form_Load()
    RefreshDashboard
End Sub

Private Sub RefreshDashboard()
    On Error Resume Next

    ' System Health
    Me.lblActiveUsers = GetActiveUserCount()
    Me.lblBackendSize = Format(GetBackendFileSizeMB(), "0.0") & " MB"
    Me.lblLastBackup = GetLastBackupDate()
    Me.lblVersion = GetCurrentVersion()

    ' Orders last 7 days
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("qryAdminDashboard_OrdersByDay", dbOpenSnapshot)
    Dim totalOrders As Long
    totalOrders = 0
    Do While Not rs.EOF
        totalOrders = totalOrders + rs!OrderCount
        rs.MoveNext
    Loop
    rs.Close
    Me.lblTotalOrders = totalOrders

    ' Top users
    Me.lstTopUsers.RowSource = "qryAdminDashboard_TopUsers"

    ' Concurrency stats
    Set rs = CurrentDb.OpenRecordset("qryAdminDashboard_ConcurrencyStats", dbOpenSnapshot)
    If Not rs.EOF Then
        Me.lblRetryCount = Nz(rs!RetryCount, 0)
        Me.lblTimeoutCount = Nz(rs!TimeoutCount, 0)
        Me.lblErrorCount = Nz(rs!ErrorCount, 0)
    End If
    rs.Close

    Me.lblLastRefresh = "Last refreshed: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub btnRefresh_Click()
    RefreshDashboard
End Sub

Private Sub btnCompactBackend_Click()
    If MsgBox("Compact backend database? Ensure no users are connected.", vbYesNo + vbQuestion) = vbYes Then
        ' Call compact function
        CompactBackendDatabase
    End If
End Sub

Private Sub btnViewLog_Click()
    DoCmd.OpenQuery "qryConcurrencyLog_Last24Hours"
End Sub

HELPER FUNCTIONS (in form module or basAdminTools):

Public Function GetActiveUserCount() As Long
    ' Read .ldb file and count users
    ' Return number of active connections
End Function

Public Function GetBackendFileSizeMB() As Double
    ' Use FileSystemObject to get backend file size
    ' Return size in megabytes
End Function

Public Function GetLastBackupDate() As String
    ' Check backup folder for most recent file
    ' Return formatted date string
End Function

Public Sub CompactBackendDatabase()
    ' Compact backend using DBEngine.CompactDatabase
    ' Show progress and result
End Sub

PROVIDE:
- SQL for all three queries
- Form design layout description
- Complete form code module
- Helper functions for system health checks
- Button click handlers for all quick actions
```

**After implementation:**
- Open frmAdminDashboard
- Verify all sections display data
- Test each button
- Use for ongoing monitoring

---

# Additional Useful Prompts

## Update Retry Count from Config

**If you want to make retry count configurable:**

```
Update basSeqAllocator.ReserveSeq to read retry count from configuration
instead of hardcoding 5 attempts.

CHANGES:
1. At start of ReserveSeq function, read retry count from config:
   Dim maxRetries As Integer
   maxRetries = GetConfigNumber("SequenceRetryAttempts")
   If maxRetries = 0 Then maxRetries = 5  ' Default fallback

2. Also read retry delay:
   Dim retryDelayMs As Long
   retryDelayMs = GetConfigNumber("SequenceRetryDelayMs")
   If retryDelayMs = 0 Then retryDelayMs = 500  ' Default fallback

3. Use these variables in the retry loop instead of hardcoded values

4. Update error message to use maxRetries variable:
   "Could not reserve sequence after " & maxRetries & " attempts..."

This allows admins to adjust retry behavior via tblConfig without code changes.
```

---

## Create Startup Version Check

**Add version check when database opens:**

```
Create startup version validation to ensure users have minimum required version.

Add to your main form's Form_Open event:

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo EH

    ' Get current and minimum versions
    Dim current As String
    Dim minRequired As String

    current = GetCurrentVersion()
    minRequired = GetConfig("MinRequiredVersion")

    ' Compare versions
    If CompareVersions(current, minRequired) < 0 Then
        MsgBox "Your database version (" & current & ") is outdated." & vbCrLf & vbCrLf & _
               "Minimum required version: " & minRequired & vbCrLf & vbCrLf & _
               "Please update your frontend copy before continuing." & vbCrLf & _
               "Contact your administrator if you need help.", _
               vbCritical, "Update Required"
        Cancel = True
        Application.Quit
        Exit Sub
    End If

    ' Optional: Check for recommended updates (non-blocking)
    If IsUpdateAvailable() Then
        If MsgBox("A new version is available. Would you like to see release notes?", _
                  vbYesNo + vbInformation, "Update Available") = vbYes Then
            DoCmd.OpenForm "frmAbout"
        End If
    End If

    Exit Sub

EH:
    ' If version check fails, allow access (don't block user)
    Debug.Print "Version check error: " & Err.Description
    Resume Next
End Sub

This ensures users stay on supported versions and alerts them to updates.
```

---

## Create Table Link Repair Function

**Fix broken table links automatically:**

```
Create a function to repair broken table links if backend path changes.

CREATE MODULE: basTableLinkManager

Public Function RepairTableLinks() As Boolean
    ' Repairs all linked tables to point to current backend

    On Error GoTo EH

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim backendPath As String
    Dim fixedCount As Long

    Set db = CurrentDb
    backendPath = GetConfig("BackendPath")

    ' Check if backend exists
    If Not Dir(backendPath) <> "" Then
        MsgBox "Backend database not found at: " & backendPath, vbCritical
        RepairTableLinks = False
        Exit Function
    End If

    fixedCount = 0

    ' Loop through all tables
    For Each tdf In db.TableDefs
        ' Check if it's a linked table
        If Len(tdf.Connect) > 0 Then
            ' Update connection string
            tdf.Connect = ";DATABASE=" & backendPath
            tdf.RefreshLink
            fixedCount = fixedCount + 1
        End If
    Next tdf

    MsgBox "Repaired " & fixedCount & " linked tables.", vbInformation
    RepairTableLinks = True

    Exit Function

EH:
    MsgBox "Error repairing links: " & Err.Description, vbCritical
    RepairTableLinks = False
End Function

Public Sub TestTableLinks()
    ' Test if all linked tables are accessible

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rs As DAO.Recordset
    Dim brokenCount As Long

    Set db = CurrentDb
    brokenCount = 0

    For Each tdf In db.TableDefs
        If Len(tdf.Connect) > 0 Then
            On Error Resume Next
            Set rs = db.OpenRecordset(tdf.Name, dbOpenSnapshot)
            If Err.Number <> 0 Then
                Debug.Print "Broken link: " & tdf.Name
                brokenCount = brokenCount + 1
            Else
                rs.Close
            End If
            On Error GoTo 0
        End If
    Next tdf

    If brokenCount > 0 Then
        MsgBox brokenCount & " broken table links found. Run RepairTableLinks to fix.", vbExclamation
    Else
        MsgBox "All table links are working.", vbInformation
    End If
End Sub

Add button to admin dashboard or tools menu to call these functions.
```

---

# Testing Checklist

After implementing with Copilot, verify:

**Phase 1 Tests:**
- [ ] Code compiles without errors
- [ ] Single user can reserve sequences
- [ ] Two users simultaneously - no duplicates
- [ ] Concurrency log shows events
- [ ] No crashes during retry logic

**Phase 2 Tests:**
- [ ] Config values read correctly
- [ ] Version comparison works (test with 1.0.0, 1.1.0, 2.0.0)
- [ ] frmAbout displays correctly
- [ ] Remote connection detection works
- [ ] Error messages are user-friendly

**Phase 3 Tests:**
- [ ] All automated tests pass
- [ ] Manual concurrent test shows no duplicates
- [ ] Dashboard displays current data
- [ ] Admin tools work

---

# Quick Reference

**Essential Test Commands:**

```vba
' Test sequence reservation
?ReserveSeq("TEST", "TST", "", "T", False)

' Test configuration
?GetConfig("BackendPath")
?GetConfigNumber("SequenceRetryAttempts")

' Test versioning
?GetCurrentVersion()
?CompareVersions("1.0.0", "1.1.0")

' Test remote access
?IsRemoteConnection()
?TestBackendConnection()

' Test logging
LogConcurrencyEvent "TEST", "Testing system"

' Run all tests
RunAllTests

' Validate sequences
?ValidateOrderSequences()
```

**Essential Queries:**

```sql
-- Check concurrency log
SELECT * FROM qryConcurrencyLog_Last24Hours;

-- Check for duplicate orders
SELECT OrderNumber, COUNT(*) AS cnt
FROM SalesOrders
GROUP BY OrderNumber
HAVING COUNT(*) > 1;

-- Check sequence status
SELECT * FROM OrderSeq ORDER BY LastUpdated DESC;
```

---

# Troubleshooting Copilot Generated Code

**If code doesn't compile:**
1. Check for missing references (Tools → References)
2. Look for typos in function names
3. Verify all required tables exist
4. Ask Copilot to fix specific error

**If tests fail:**
1. Run Debug → Compile to find syntax errors
2. Check if all dependencies are in place
3. Review error messages carefully
4. Test components individually

**If Copilot's solution isn't quite right:**
1. Be more specific in next prompt
2. Show Copilot the error message
3. Ask for modifications to specific parts
4. Reference existing code structure

---

# Success Indicators

**You know it's working when:**

✅ Code compiles without errors
✅ Tests pass (especially no duplicates!)
✅ Concurrency log shows activity
✅ Forms open and display data
✅ Multiple users can work simultaneously
✅ Error messages are clear and helpful

**Time to deploy when:**

✅ All Phase 1 complete and tested
✅ Most of Phase 2 complete
✅ Multi-user test passed with 2-3 users
✅ Documentation created
✅ Backups made
✅ Users notified

---

**Good luck with your implementation! 🚀**

Remember: Do Phase 1 first - it's the most critical. Everything else is enhancement.

Attribute VB_Name = "basResetForCleanCopy"
Option Compare Database
Option Explicit

' ============================================================================
' basResetForCleanCopy
' ----------------------------------------------------------------------------
' Clears all transactional tables so the backend can be distributed as a
' clean starting copy.  Lookup/reference tables (QualifierType, SystemLetter,
' etc.) are left untouched.
'
' Tables cleared (FK-safe order):
'   1. SalesOrderEntry   - child of SalesOrders via SOID
'   2. tblOrderAudit     - child of SalesOrders via SOID
'   3. tblConcurrencyLog - independent log table
'   4. SalesOrders       - parent order table
'   5. OrderSeq          - sequence counter table
'
' HOW TO RUN:
'   Open the Immediate Window (Ctrl+G) and type:
'       ResetAllTablesForCleanCopy
' ============================================================================

Public Sub ResetAllTablesForCleanCopy()

    ' -----------------------------------------------------------------------
    ' Safety confirmation -- two prompts required
    ' -----------------------------------------------------------------------
    Dim msg1 As String
    msg1 = "WARNING: This will permanently delete ALL records from:" & vbCrLf & vbCrLf & _
           "  - SalesOrderEntry" & vbCrLf & _
           "  - tblOrderAudit" & vbCrLf & _
           "  - tblConcurrencyLog" & vbCrLf & _
           "  - SalesOrders" & vbCrLf & _
           "  - OrderSeq" & vbCrLf & vbCrLf & _
           "Reference/lookup tables will NOT be touched." & vbCrLf & vbCrLf & _
           "Are you sure you want to continue?"

    If MsgBox(msg1, vbExclamation + vbYesNo, "Reset Database -- Step 1 of 2") <> vbYes Then
        MsgBox "Reset cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    Dim msg2 As String
    msg2 = "SECOND CONFIRMATION required." & vbCrLf & vbCrLf & _
           "This action CANNOT be undone." & vbCrLf & _
           "Make a backup first if you need to preserve existing data." & vbCrLf & vbCrLf & _
           "Proceed with deleting all transactional records?"

    If MsgBox(msg2, vbCritical + vbYesNo, "Reset Database -- Step 2 of 2") <> vbYes Then
        MsgBox "Reset cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' Capture before-counts for the summary
    ' -----------------------------------------------------------------------
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim cntEntry As Long, cntAudit As Long, cntConc As Long
    Dim cntOrders As Long, cntSeq As Long

    On Error Resume Next
    cntEntry  = DCount("*", "SalesOrderEntry")
    cntAudit  = DCount("*", "tblOrderAudit")
    cntConc   = DCount("*", "tblConcurrencyLog")
    cntOrders = DCount("*", "SalesOrders")
    cntSeq    = DCount("*", "OrderSeq")
    On Error GoTo EH

    ' -----------------------------------------------------------------------
    ' Delete in FK-safe order
    ' -----------------------------------------------------------------------
    Debug.Print "--- ResetAllTablesForCleanCopy: " & Now() & " ---"

    ' 1. Children first
    db.Execute "DELETE FROM SalesOrderEntry;", dbFailOnError
    Debug.Print "  Cleared SalesOrderEntry  (" & cntEntry & " rows)"

    db.Execute "DELETE FROM tblOrderAudit;", dbFailOnError
    Debug.Print "  Cleared tblOrderAudit    (" & cntAudit & " rows)"

    db.Execute "DELETE FROM tblConcurrencyLog;", dbFailOnError
    Debug.Print "  Cleared tblConcurrencyLog(" & cntConc & " rows)"

    ' 2. Parent
    db.Execute "DELETE FROM SalesOrders;", dbFailOnError
    Debug.Print "  Cleared SalesOrders      (" & cntOrders & " rows)"

    ' 3. Sequence table
    db.Execute "DELETE FROM OrderSeq;", dbFailOnError
    Debug.Print "  Cleared OrderSeq         (" & cntSeq & " rows)"

    ' -----------------------------------------------------------------------
    ' Summary
    ' -----------------------------------------------------------------------
    Dim summary As String
    summary = "Reset complete.  Records removed:" & vbCrLf & vbCrLf & _
              "  SalesOrderEntry   : " & cntEntry  & vbCrLf & _
              "  tblOrderAudit     : " & cntAudit  & vbCrLf & _
              "  tblConcurrencyLog : " & cntConc   & vbCrLf & _
              "  SalesOrders       : " & cntOrders & vbCrLf & _
              "  OrderSeq          : " & cntSeq    & vbCrLf & vbCrLf & _
              "The database is ready to distribute as a clean copy." & vbCrLf & _
              "Lookup tables (QualifierType, SystemLetter, etc.) were NOT changed."

    MsgBox summary, vbInformation, "Reset Complete"
    Debug.Print "  Done."
    Set db = Nothing
    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum  = Err.Number
    errDesc = Err.Description
    Set db  = Nothing
    MsgBox "Error " & errNum & " during reset:" & vbCrLf & errDesc & vbCrLf & vbCrLf & _
           "Some tables may have been partially cleared.  Check the Immediate Window for progress.", _
           vbCritical, "Reset Error"
    Debug.Print "  ERROR " & errNum & ": " & errDesc
End Sub

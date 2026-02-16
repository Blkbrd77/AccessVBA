'==============================================================
' Module: basRelinkTables
'
' PURPOSE:
'   Re-points all linked (BE) tables in the current FE to the
'   path stored in tblConfig("BackendPath"), then calls
'   RefreshLink on each one.
'
'   Call this once after:
'     - The BE file is moved or renamed
'     - A new FE copy is distributed that still has the old
'       split-time path baked into its TableDef.Connect strings
'     - tblConfig.BackendPath is updated to a new location
'
' DEPENDENCIES:
'   - basConfig : GetConfig("BackendPath")
'
' PUBLIC INTERFACE:
'   RelinkAllTables()   As Boolean   ' main entry point
'   GetLinkedBackends() As String    ' diagnostic -- lists current paths
'==============================================================
Option Compare Database
Option Explicit

'==============================================================
' PUBLIC: RelinkAllTables
'
' Reads BackendPath from tblConfig, then updates every linked
' TableDef's Connect string and calls RefreshLink.
'
' Returns True  if all tables were re-linked successfully.
' Returns False if BackendPath is blank, the file is missing,
'               or any individual RefreshLink call fails.
'
' On failure a plain-language MsgBox is shown; the caller does
' not need to display its own error.
'==============================================================
Public Function RelinkAllTables() As Boolean

    On Error GoTo EH

    Dim newPath  As String
    Dim db       As DAO.Database
    Dim tdf      As DAO.TableDef
    Dim linked   As Long
    Dim failed   As Long
    Dim msg      As String

    ' ---- 1. Resolve target path -----------------------------------
    newPath = Nz(GetConfig("BackendPath"), vbNullString)
    newPath = Trim$(newPath)

    If Len(newPath) = 0 Then
        MsgBox "BackendPath is not set in tblConfig." & vbCrLf & _
               "Open tblConfig and add a row with ConfigKey = ""BackendPath"" " & _
               "and ConfigValue = the full path to the BE .accdb file.", _
               vbExclamation, "Relink Failed"
        RelinkAllTables = False
        Exit Function
    End If

    ' ---- 2. Confirm the file actually exists ----------------------
    If Dir(newPath) = "" Then
        MsgBox "Backend file not found:" & vbCrLf & vbCrLf & _
               newPath & vbCrLf & vbCrLf & _
               "Check tblConfig.BackendPath and try again.", _
               vbExclamation, "Relink Failed"
        RelinkAllTables = False
        Exit Function
    End If

    ' ---- 3. Re-link every linked table ----------------------------
    Set db = CurrentDb

    For Each tdf In db.TableDefs

        ' Skip local, system, and temp tables
        If Len(tdf.Connect & "") > 0 Then

            linked = linked + 1

            On Error Resume Next

            tdf.Connect = ";DATABASE=" & newPath
            tdf.RefreshLink

            If Err.Number <> 0 Then
                failed = failed + 1
                Debug.Print "RelinkAllTables: FAILED on [" & tdf.Name & "] " & _
                            "Err " & Err.Number & ": " & Err.Description
                Err.Clear
            Else
                Debug.Print "RelinkAllTables: OK  [" & tdf.Name & "]"
            End If

            On Error GoTo EH

        End If

    Next tdf

    ' ---- 4. Report result -----------------------------------------
    If linked = 0 Then
        MsgBox "No linked tables were found in this frontend." & vbCrLf & _
               "Nothing was changed.", _
               vbInformation, "Relink"
        RelinkAllTables = True
        GoTo CleanExit
    End If

    If failed = 0 Then
        MsgBox "All " & linked & " linked table(s) successfully re-pointed to:" & _
               vbCrLf & vbCrLf & newPath, _
               vbInformation, "Relink Complete"
        RelinkAllTables = True
    Else
        msg = failed & " of " & linked & " table(s) failed to relink." & vbCrLf & vbCrLf & _
              "Check the Immediate Window (Ctrl+G) for details." & vbCrLf & vbCrLf & _
              "Target path was:" & vbCrLf & newPath
        MsgBox msg, vbExclamation, "Relink Partially Failed"
        RelinkAllTables = False
    End If

CleanExit:
    Set tdf = Nothing
    Set db  = Nothing
    Exit Function

EH:
    MsgBox "Unexpected error in RelinkAllTables" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Relink Error"
    RelinkAllTables = False
    Resume CleanExit

End Function

'==============================================================
' PUBLIC: GetLinkedBackends
'
' Diagnostic helper. Returns a string listing every distinct
' backend path currently baked into the FE's linked tables,
' plus a count of how many tables point to each one.
'
' Usage: Debug.Print GetLinkedBackends()
'==============================================================
Public Function GetLinkedBackends() As String

    On Error GoTo EH

    Dim tdf   As DAO.TableDef
    Dim dict  As Object          ' Scripting.Dictionary
    Dim key   As Variant
    Dim conn  As String
    Dim s     As String

    Set dict = CreateObject("Scripting.Dictionary")

    For Each tdf In CurrentDb.TableDefs
        If Len(tdf.Connect & "") > 0 Then
            conn = tdf.Connect
            If dict.Exists(conn) Then
                dict(conn) = dict(conn) + 1
            Else
                dict.Add conn, 1
            End If
        End If
    Next tdf

    If dict.Count = 0 Then
        GetLinkedBackends = "(no linked tables found)"
        GoTo CleanExit
    End If

    s = "Linked backends in this FE:" & vbCrLf
    For Each key In dict.Keys
        s = s & "  [" & dict(key) & " table(s)]  " & key & vbCrLf
    Next key

    GetLinkedBackends = s

CleanExit:
    Set dict = Nothing
    Set tdf  = Nothing
    Exit Function

EH:
    GetLinkedBackends = "Error " & Err.Number & ": " & Err.Description
    Resume CleanExit

End Function

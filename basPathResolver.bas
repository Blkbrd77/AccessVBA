'==============================================================
' Module: basPathResolver
'
' PURPOSE:
'   Resolves the correct backend (.accdb) path at startup by
'   trying a prioritised list of candidate paths from tblConfig.
'
'   This solves the VPN / mapped-drive problem: office users
'   may reach the backend as S:\Data\BE.accdb while VPN users
'   reach it as \\server\share\Data\BE.accdb (or a different
'   drive letter altogether). Both paths point to the same
'   physical file; we just need to find the first one that
'   Windows can resolve right now.
'
' HOW IT WORKS:
'   1. Read every tblConfig row whose ConfigKey starts with
'      "BackendPath" (e.g. BackendPath, BackendPath_VPN,
'      BackendPath_2).  Keys are sorted alphabetically so
'      "BackendPath" (plain) is always tried first.
'   2. Test each path with Dir() -- no network dialog, instant.
'   3. The first path that resolves is written back to
'      "BackendPath" in tblConfig so that RelinkAllTables (and
'      anything else that calls GetConfig("BackendPath")) picks
'      up the correct value for this session.
'   4. Call RelinkAllTables to reconnect all linked tables.
'
' SETUP (one-time):
'   Add rows to tblConfig:
'     ConfigKey = "BackendPath"      ConfigValue = "S:\Data\BE.accdb"
'     ConfigKey = "BackendPath_VPN"  ConfigValue = "\\server\share\Data\BE.accdb"
'
'   You can add as many BackendPath_* variants as you need.
'   The plain "BackendPath" key is the primary (LAN) path and is
'   always tried first because it sorts before "BackendPath_".
'
' PUBLIC INTERFACE:
'   ResolveAndRelinkBackend()  As Boolean   ' call from AutoExec / Form_Load
'   GetResolvedBackendPath()   As String    ' returns last resolved path
'==============================================================
Option Compare Database
Option Explicit

Private m_ResolvedPath As String   ' cached after successful resolve

'==============================================================
' PUBLIC: ResolveAndRelinkBackend
'
' Tries each BackendPath* candidate in tblConfig until one
' resolves, writes it to "BackendPath", then relinks all tables.
'
' Returns True  if a reachable path was found and all tables
'               were relinked successfully.
' Returns False if no candidate path was reachable, or if
'               RelinkAllTables reported a failure.
'
' On failure a plain-language MsgBox is shown.
'==============================================================
Public Function ResolveAndRelinkBackend() As Boolean

    On Error GoTo EH

    Dim candidates() As String
    Dim n            As Long
    Dim i            As Long
    Dim resolved     As String

    ' ---- 1. Collect all BackendPath* values from tblConfig ----
    n = GetBackendCandidates(candidates)

    If n = 0 Then
        MsgBox "No BackendPath entries found in tblConfig." & vbCrLf & vbCrLf & _
               "Add at least one row:" & vbCrLf & _
               "  ConfigKey = ""BackendPath""" & vbCrLf & _
               "  ConfigValue = full path to the BE .accdb file.", _
               vbExclamation, "Backend Not Configured"
        ResolveAndRelinkBackend = False
        Exit Function
    End If

    ' ---- 2. Try each candidate --------------------------------
    resolved = vbNullString
    For i = 0 To n - 1
        If Len(Trim$(candidates(i))) > 0 Then
            If Dir(Trim$(candidates(i))) <> vbNullString Then
                resolved = Trim$(candidates(i))
                Exit For
            End If
        End If
    Next i

    ' ---- 3. Nothing worked ------------------------------------
    If Len(resolved) = 0 Then
        Dim listMsg As String
        listMsg = "Tried " & n & " path(s), none were reachable:" & vbCrLf
        For i = 0 To n - 1
            listMsg = listMsg & "  " & candidates(i) & vbCrLf
        Next i
        listMsg = listMsg & vbCrLf & _
                  "If you are on VPN, ensure it is connected." & vbCrLf & _
                  "If you are in the office, check your drive mapping." & vbCrLf & _
                  "Add a BackendPath_* row in tblConfig for any missing path."
        MsgBox listMsg, vbExclamation, "Backend Unreachable"
        ResolveAndRelinkBackend = False
        Exit Function
    End If

    ' ---- 4. Persist resolved path so RelinkAllTables uses it --
    m_ResolvedPath = resolved
    SetConfig "BackendPath", resolved

    ' ---- 5. Relink all linked tables to the resolved path -----
    ResolveAndRelinkBackend = RelinkAllTables()

    Exit Function

EH:
    MsgBox "Unexpected error in ResolveAndRelinkBackend" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Path Resolver Error"
    ResolveAndRelinkBackend = False

End Function

'==============================================================
' PUBLIC: GetResolvedBackendPath
'
' Returns the path that was successfully resolved during the
' last call to ResolveAndRelinkBackend, or an empty string if
' it has not been called yet this session.
'==============================================================
Public Function GetResolvedBackendPath() As String
    GetResolvedBackendPath = m_ResolvedPath
End Function

'==============================================================
' PRIVATE: GetBackendCandidates
'
' Queries tblConfig for every row whose ConfigKey starts with
' "BackendPath" (case-insensitive) and returns the values as a
' 0-based String array sorted by ConfigKey alphabetically.
'
' The plain key "BackendPath" sorts before "BackendPath_VPN"
' etc., so the primary (LAN) path is always tried first.
'
' Returns the number of candidates found (0 if none).
'==============================================================
Private Function GetBackendCandidates(ByRef out() As String) As Long

    On Error GoTo EH

    Dim rs    As DAO.Recordset
    Dim db    As DAO.Database
    Dim sql   As String
    Dim count As Long

    sql = "SELECT ConfigValue FROM tblConfig " & _
          "WHERE ConfigKey Like 'BackendPath*' " & _
          "ORDER BY ConfigKey;"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    count = 0
    ReDim out(0)

    Do While Not rs.EOF
        If Not IsNull(rs!ConfigValue) Then
            If count = 0 Then
                ReDim out(0)
            Else
                ReDim Preserve out(count)
            End If
            out(count) = CStr(rs!ConfigValue)
            count = count + 1
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetBackendCandidates = count
    Exit Function

EH:
    GetBackendCandidates = 0
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing

End Function

'==============================================================
' PRIVATE: SetConfig
'
' Writes a value back to tblConfig. Used to persist the
' resolved path into "BackendPath" so downstream code that
' calls GetConfig("BackendPath") picks up the correct value.
'
' If the key already exists it is updated; otherwise inserted.
'==============================================================
Private Sub SetConfig(ByVal configKey As String, ByVal configValue As String)

    On Error GoTo EH

    Dim db  As DAO.Database
    Dim rs  As DAO.Recordset
    Dim sql As String

    Set db = CurrentDb

    sql = "SELECT ConfigValue FROM tblConfig WHERE ConfigKey = '" & configKey & "';"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        ' Row does not exist -- insert it
        rs.AddNew
        rs!ConfigKey   = configKey
        rs!ConfigValue = configValue
        rs.Update
    Else
        ' Row exists -- update it
        rs.Edit
        rs!ConfigValue = configValue
        rs.Update
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    ' Force CurrentDb to see the fresh value next time GetConfig is called
    CurrentDb.QueryDefs.Refresh

    Exit Sub

EH:
    ' Non-fatal: log to Immediate window but don't crash the caller
    Debug.Print "SetConfig error: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing

End Sub

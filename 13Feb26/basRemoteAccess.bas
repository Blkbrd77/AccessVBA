'==============================================================
' Module: basRemoteAccess
'
' PURPOSE:
'   Provides backend connectivity testing and user-friendly
'   error handling for a split FE/BE Access database.
'
'   This module does NOT attempt to detect VPN or classify
'   users as remote vs local. The backend UNC path either
'   resolves or it doesn't -- Access will surface the error
'   naturally. This module catches those errors and presents
'   them in plain language with Retry/Cancel.
'
' DEPENDENCIES:
'   - basConfig    : GetConfig("BackendPath")
'
' NOTE: No dependency on basAuditLog. LogConcurrencyEvent is NOT
' called here because tblConcurrencyLog is a linked BE table --
' logging to it during a connection test is circular (the BE may
' be unreachable, which is exactly what we are testing for).
'
' PUBLIC INTERFACE:
'   TestBackendConnection()          As Boolean
'   HandleNetworkError(...)          As Boolean
'   GetLastBackendErrorNumber()      As Long
'   GetLastBackendErrorDescription() As String
'   GetLastBackendLatencyMs()        As Long
'
' COMPILE-TIME CHANGE REQUIRED IN CALLERS:
'   Form_frmOrderList.Form_Load -- delete the IsRemoteConnection()
'   block (the 6-line If/MsgBox block). Everything else in that
'   form (EnsureBackendConnectionOrExit, the EH calling
'   HandleNetworkError) is unchanged and works with this module.
'==============================================================
Option Compare Database
Option Explicit

'--------------------------------------------------------------
' Module-level state set by TestBackendConnection
'--------------------------------------------------------------
Private m_LastErrNumber      As Long
Private m_LastErrDescription As String
Private m_LastLatencyMs      As Long

'==============================================================
' PUBLIC: TestBackendConnection
'
' Opens the backend database read-only via ADO to confirm it
' is reachable. Stores the result in module-level variables
' so the caller can retrieve error detail via the Get* functions.
'
' Returns True  if the backend opened successfully.
' Returns False if the path is blank, unreachable, or errors.
'==============================================================
Public Function TestBackendConnection() As Boolean

    On Error GoTo EH

    Dim backendPath As String
    Dim errText     As String
    Dim t0          As Double
    Dim t1          As Double
    Dim ms          As Long
    Dim ok          As Boolean

    m_LastErrNumber      = 0
    m_LastErrDescription = vbNullString
    m_LastLatencyMs      = 0

    backendPath = Nz(GetConfig("BackendPath"), vbNullString)

    If Len(Trim$(backendPath)) = 0 Then
        m_LastErrNumber      = vbObjectError + 9001
        m_LastErrDescription = "BackendPath is not configured. Open tblConfig and set the BackendPath key."
        TestBackendConnection = False
        Exit Function
    End If

    t0 = Timer
    ok = OpenBackendReadOnly(backendPath, 5, errText)
    t1 = Timer

    ms              = ElapsedMs(t0, t1)
    m_LastLatencyMs = ms

    If ok Then
        TestBackendConnection = True
    Else
        m_LastErrNumber      = vbObjectError + 9002
        m_LastErrDescription = errText
        TestBackendConnection = False
    End If

    Exit Function

EH:
    m_LastErrNumber      = Err.Number
    m_LastErrDescription = Err.Description
    TestBackendConnection = False

End Function

'==============================================================
' PUBLIC: HandleNetworkError
'
' Displays a plain-language error message when the backend
' cannot be reached. Offers Retry and Cancel.
'
' Returns True  if the user clicked Retry.
' Returns False if the user clicked Cancel.
'
' The caller is responsible for acting on the return value
' (retry the connection or quit the application).
'==============================================================
Public Function HandleNetworkError( _
    ByVal ErrorNumber      As Long, _
    ByVal ErrorDescription As String _
) As Boolean

    On Error GoTo EH

    Dim msg  As String
    Dim resp As VbMsgBoxResult

    msg = "Cannot connect to the database." & vbCrLf & vbCrLf & _
          "This usually means:" & vbCrLf & _
          "  - The server is unreachable (check your network connection)" & vbCrLf & _
          "  - VPN is not connected (remote users must connect VPN first)" & vbCrLf & _
          "  - The shared drive is temporarily unavailable" & vbCrLf & vbCrLf & _
          "Click Retry to try again, or Cancel to exit." & vbCrLf & vbCrLf & _
          "Error " & ErrorNumber & ": " & ErrorDescription

    resp = MsgBox(msg, vbExclamation Or vbRetryCancel, "Database Connection Issue")

    HandleNetworkError = (resp = vbRetry)
    Exit Function

EH:
    HandleNetworkError = False

End Function

'==============================================================
' PUBLIC: State accessors
' Call these after TestBackendConnection returns False to get
' the error detail for passing to HandleNetworkError.
'==============================================================
Public Function GetLastBackendErrorNumber() As Long
    GetLastBackendErrorNumber = m_LastErrNumber
End Function

Public Function GetLastBackendErrorDescription() As String
    GetLastBackendErrorDescription = m_LastErrDescription
End Function

Public Function GetLastBackendLatencyMs() As Long
    GetLastBackendLatencyMs = m_LastLatencyMs
End Function

'==============================================================
' PRIVATE: OpenBackendReadOnly
'
' Attempts to open the backend .accdb via ADO (late-bound).
' Using ADO avoids triggering Access's own "file not found"
' dialogs during the connectivity test.
'
' Parameters:
'   backendPath    - Full UNC or local path to the .accdb file
'   timeoutSeconds - Connection timeout in seconds (recommend 5)
'   errText        - OUT: error description if failed
'
' Returns True if connection opened and closed cleanly.
'==============================================================
Private Function OpenBackendReadOnly( _
    ByVal backendPath    As String, _
    ByVal timeoutSeconds As Long, _
    ByRef errText        As String _
) As Boolean

    On Error GoTo EH

    Dim cn As Object  ' ADODB.Connection -- late-bound, no extra reference needed

    errText             = vbNullString
    OpenBackendReadOnly = False

    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = CLng(timeoutSeconds)
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & backendPath & ";" & _
            "Mode=Read;" & _
            "Persist Security Info=False;"
    cn.Close
    Set cn = Nothing

    OpenBackendReadOnly = True
    Exit Function

EH:
    m_LastErrNumber      = Err.Number
    m_LastErrDescription = Err.Description
    errText              = Err.Description

    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
    End If
    Set cn = Nothing

End Function

'==============================================================
' PRIVATE: ElapsedMs
'
' Calculates elapsed milliseconds between two Timer values.
' Handles midnight rollover (Timer resets to 0 at midnight).
'==============================================================
Private Function ElapsedMs(ByVal t0 As Double, ByVal t1 As Double) As Long
    Dim dt As Double
    dt = t1 - t0
    If dt < 0 Then dt = dt + 86400#  ' midnight rollover
    ElapsedMs = CLng(dt * 1000#)
End Function

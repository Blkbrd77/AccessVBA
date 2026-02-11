Attribute VB_Name = "modAI_ContextExport"
Option Compare Database
Option Explicit

' ============================================================
' AI Context Export for Microsoft Access
' Creates a comprehensive context file for AI assistants
' (Copilot, Claude, ChatGPT, etc.)
'
' Entry point:
'   Call ExportAIContext
'   Call ExportAIContext "C:\MyFolder"  ' custom output path
' ============================================================

Public Sub ExportAIContext(Optional ByVal OutputFolder As String = "")
    On Error GoTo ErrHandler

    ' Default to database folder if not specified
    If Len(OutputFolder) = 0 Then
        OutputFolder = CurrentProject.Path
    End If

    Dim report As String, ts As String, outFile As String
    ts = Format(Now, "yyyy-mm-dd_hhnnss")
    outFile = BuildPath(OutputFolder, "AI_Context_" & GetDbNameOnly() & "_" & ts & ".txt")

    ' Build the context document
    report = BuildAISummaryHeader()
    report = report & BuildEnvironmentSection()
    report = report & BuildTableSchemaSection()
    report = report & BuildRelationshipsSection()
    report = report & BuildQueriesSection()
    report = report & BuildAllFormsSection()
    report = report & BuildAllReportsSection()
    report = report & BuildAllCodeModulesSection()

    ' Write output
    EnsureFolder OutputFolder
    WriteFile outFile, report

    MsgBox "AI Context exported to:" & vbCrLf & vbCrLf & outFile, vbInformation, "Export Complete"
    Exit Sub

ErrHandler:
    MsgBox "ExportAIContext failed:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ============================================================
' Summary Header - Quick overview for AI
' ============================================================

Private Function BuildAISummaryHeader() As String
    Dim s As String
    Dim tblCount As Long, qryCount As Long, frmCount As Long, rptCount As Long, modCount As Long

    tblCount = GetUserTableCount()
    qryCount = CurrentData.AllQueries.Count
    frmCount = CurrentProject.AllForms.Count
    rptCount = CurrentProject.AllReports.Count
    modCount = GetCodeModuleCount()

    s = "================================================================" & vbCrLf
    s = s & "AI CONTEXT PACKAGE - MICROSOFT ACCESS DATABASE" & vbCrLf
    s = s & "================================================================" & vbCrLf
    s = s & "Generated: " & Now & vbCrLf
    s = s & "Database: " & CurrentProject.Name & vbCrLf
    s = s & "Path: " & CurrentDb.Name & vbCrLf
    s = s & vbCrLf
    s = s & "QUICK STATS:" & vbCrLf
    s = s & "  Tables: " & tblCount & vbCrLf
    s = s & "  Queries: " & qryCount & vbCrLf
    s = s & "  Forms: " & frmCount & vbCrLf
    s = s & "  Reports: " & rptCount & vbCrLf
    s = s & "  Code Modules: " & modCount & vbCrLf
    s = s & vbCrLf
    s = s & "This file contains the complete structure and code of this" & vbCrLf
    s = s & "Access database for AI assistant context." & vbCrLf
    s = s & "================================================================" & vbCrLf & vbCrLf

    BuildAISummaryHeader = s
End Function

' ============================================================
' Environment Section
' ============================================================

Private Function BuildEnvironmentSection() As String
    Dim s As String
    Dim linkedCount As Long, localCount As Long

    s = SectionHeader("ENVIRONMENT")
    s = s & "Access Version: " & Application.Version & " (" & GetBitness() & ")" & vbCrLf
    s = s & "Windows: " & GetWindowsVersion() & vbCrLf
    s = s & vbCrLf

    ' Split database info
    GetTableCounts linkedCount, localCount
    s = s & "Database Type: " & IIf(linkedCount > 0, "Split (Front-end/Back-end)", "Standalone") & vbCrLf
    s = s & "Local Tables: " & localCount & vbCrLf
    s = s & "Linked Tables: " & linkedCount & vbCrLf

    If linkedCount > 0 Then
        s = s & vbCrLf & "Back-end Source(s):" & vbCrLf
        s = s & GetBackendList()
    End If

    s = s & vbCrLf & "VBA References:" & vbCrLf
    s = s & GetReferences() & vbCrLf

    BuildEnvironmentSection = s
End Function

' ============================================================
' Table Schema Section - Field definitions, types, indexes
' ============================================================

Private Function BuildTableSchemaSection() As String
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim s As String, rowCount As Long

    s = SectionHeader("TABLE SCHEMA")

    For Each tdf In db.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 1) <> "~" Then
            s = s & "----------------------------------------" & vbCrLf
            s = s & "TABLE: " & tdf.Name

            ' Get row count
            On Error Resume Next
            rowCount = DCount("*", tdf.Name)
            If Err.Number = 0 Then
                s = s & " (" & rowCount & " rows)"
            End If
            On Error GoTo 0

            ' Show if linked
            If Len(tdf.Connect & "") > 0 Then
                s = s & " [LINKED]"
            End If
            s = s & vbCrLf & vbCrLf

            ' Fields
            s = s & "  Fields:" & vbCrLf
            For Each fld In tdf.Fields
                s = s & "    - " & fld.Name & " : " & GetFieldTypeName(fld.Type)
                If fld.Size > 0 And fld.Type = dbText Then
                    s = s & "(" & fld.Size & ")"
                End If
                If fld.Required Then s = s & " NOT NULL"
                If Len(fld.DefaultValue & "") > 0 Then s = s & " DEFAULT " & fld.DefaultValue
                s = s & vbCrLf
            Next

            ' Primary Key
            s = s & vbCrLf & "  Primary Key: "
            On Error Resume Next
            Dim pk As DAO.Index
            Set pk = Nothing
            For Each idx In tdf.Indexes
                If idx.Primary Then
                    Set pk = idx
                    Exit For
                End If
            Next
            If pk Is Nothing Then
                s = s & "(none)" & vbCrLf
            Else
                s = s & GetIndexFields(pk) & vbCrLf
            End If
            On Error GoTo 0

            ' Other Indexes
            s = s & "  Indexes:" & vbCrLf
            Dim hasIdx As Boolean: hasIdx = False
            For Each idx In tdf.Indexes
                If Not idx.Primary Then
                    hasIdx = True
                    s = s & "    - " & idx.Name & ": " & GetIndexFields(idx)
                    If idx.Unique Then s = s & " [UNIQUE]"
                    s = s & vbCrLf
                End If
            Next
            If Not hasIdx Then s = s & "    (none)" & vbCrLf

            s = s & vbCrLf
        End If
    Next

    BuildTableSchemaSection = s
End Function

' ============================================================
' Relationships Section
' ============================================================

Private Function BuildRelationshipsSection() As String
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim s As String, attrs As String, fields As String

    s = SectionHeader("RELATIONSHIPS")

    If db.Relations.Count = 0 Then
        s = s & "(No relationships defined)" & vbCrLf & vbCrLf
        BuildRelationshipsSection = s
        Exit Function
    End If

    For Each rel In db.Relations
        s = s & rel.Table & " -> " & rel.ForeignTable & vbCrLf

        ' Fields
        fields = "  Fields: "
        For Each fld In rel.Fields
            fields = fields & fld.Name & " -> " & fld.ForeignName & ", "
        Next
        If Len(fields) > 10 Then fields = Left$(fields, Len(fields) - 2)
        s = s & fields & vbCrLf

        ' Attributes
        s = s & "  Attributes: "
        If (rel.Attributes And dbRelationDontEnforce) = 0 Then
            s = s & "Enforced"
            If (rel.Attributes And dbRelationUpdateCascade) <> 0 Then s = s & ", Cascade Update"
            If (rel.Attributes And dbRelationDeleteCascade) <> 0 Then s = s & ", Cascade Delete"
        Else
            s = s & "Not Enforced"
        End If
        s = s & vbCrLf & vbCrLf
    Next

    BuildRelationshipsSection = s
End Function

' ============================================================
' Queries Section - All query SQL
' ============================================================

Private Function BuildQueriesSection() As String
    Dim db As DAO.Database: Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim s As String

    s = SectionHeader("QUERIES")

    If db.QueryDefs.Count = 0 Then
        s = s & "(No queries defined)" & vbCrLf & vbCrLf
        BuildQueriesSection = s
        Exit Function
    End If

    For Each qdf In db.QueryDefs
        ' Skip system/temp queries
        If Left$(qdf.Name, 1) <> "~" Then
            s = s & "--- QUERY: " & qdf.Name & " ---" & vbCrLf
            s = s & "Type: " & GetQueryTypeName(qdf.Type) & vbCrLf
            s = s & "SQL:" & vbCrLf
            s = s & qdf.SQL & vbCrLf & vbCrLf
        End If
    Next

    BuildQueriesSection = s
End Function

' ============================================================
' Forms Section - All forms via SaveAsText
' ============================================================

Private Function BuildAllFormsSection() As String
    Dim frm As AccessObject
    Dim s As String

    s = SectionHeader("FORMS")

    If CurrentProject.AllForms.Count = 0 Then
        s = s & "(No forms in database)" & vbCrLf & vbCrLf
        BuildAllFormsSection = s
        Exit Function
    End If

    For Each frm In CurrentProject.AllForms
        s = s & "=== FORM: " & frm.Name & " ===" & vbCrLf
        s = s & GetSaveAsTextContent(acForm, frm.Name) & vbCrLf & vbCrLf
    Next

    BuildAllFormsSection = s
End Function

' ============================================================
' Reports Section - All reports via SaveAsText
' ============================================================

Private Function BuildAllReportsSection() As String
    Dim rpt As AccessObject
    Dim s As String

    s = SectionHeader("REPORTS")

    If CurrentProject.AllReports.Count = 0 Then
        s = s & "(No reports in database)" & vbCrLf & vbCrLf
        BuildAllReportsSection = s
        Exit Function
    End If

    For Each rpt In CurrentProject.AllReports
        s = s & "=== REPORT: " & rpt.Name & " ===" & vbCrLf
        s = s & GetSaveAsTextContent(acReport, rpt.Name) & vbCrLf & vbCrLf
    Next

    BuildAllReportsSection = s
End Function

' ============================================================
' Code Modules Section - All VBA code
' ============================================================

Private Function BuildAllCodeModulesSection() As String
    Dim s As String

    s = SectionHeader("VBA CODE MODULES")

    If Not IsVBEAccessible() Then
        s = s & "*** VBA PROJECT ACCESS IS DISABLED ***" & vbCrLf
        s = s & "To include code modules, enable:" & vbCrLf
        s = s & "  File > Options > Trust Center > Trust Center Settings" & vbCrLf
        s = s & "  > Macro Settings > Trust access to the VBA project object model" & vbCrLf & vbCrLf
        BuildAllCodeModulesSection = s
        Exit Function
    End If

    On Error GoTo ErrHandler

    Dim vbProj As Object, vbComp As Object
    Dim content As String, tmp As String, ext As String

    Set vbProj = Application.VBE.ActiveVBProject

    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case 1: ext = "Standard Module"
            Case 2: ext = "Class Module"
            Case 3: ext = "UserForm"
            Case 100: ext = "Document Module"
            Case Else: ext = "Unknown"
        End Select

        s = s & "=== MODULE: " & vbComp.Name & " (" & ext & ") ===" & vbCrLf

        ' Get code directly from CodeModule
        If vbComp.CodeModule.CountOfLines > 0 Then
            s = s & vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines) & vbCrLf
        Else
            s = s & "(empty)" & vbCrLf
        End If
        s = s & vbCrLf
    Next

    BuildAllCodeModulesSection = s
    Exit Function

ErrHandler:
    s = s & "Error exporting modules: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf
    BuildAllCodeModulesSection = s
End Function

' ============================================================
' Utility Functions
' ============================================================

Private Function SectionHeader(ByVal title As String) As String
    SectionHeader = String(64, "=") & vbCrLf & _
                    title & vbCrLf & _
                    String(64, "=") & vbCrLf & vbCrLf
End Function

Private Function GetDbNameOnly() As String
    Dim n As String
    n = CurrentProject.Name
    If InStr(n, ".") > 0 Then n = Left$(n, InStrRev(n, ".") - 1)
    GetDbNameOnly = n
End Function

Private Function GetBitness() As String
    #If Win64 Then
        GetBitness = "64-bit"
    #Else
        GetBitness = "32-bit"
    #End If
End Function

Private Function GetWindowsVersion() As String
    On Error Resume Next
    Dim wmi As Object, os As Object, item As Object
    Set wmi = GetObject("winmgmts:\\.\root\CIMV2")
    Set os = wmi.ExecQuery("SELECT Caption, Version FROM Win32_OperatingSystem")
    For Each item In os
        GetWindowsVersion = Trim(item.Caption) & " " & item.Version
        Exit Function
    Next
    GetWindowsVersion = Environ$("OS")
End Function

Private Function GetUserTableCount() As Long
    Dim tdf As DAO.TableDef, cnt As Long
    For Each tdf In CurrentDb.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 1) <> "~" Then
            cnt = cnt + 1
        End If
    Next
    GetUserTableCount = cnt
End Function

Private Function GetCodeModuleCount() As Long
    On Error Resume Next
    If IsVBEAccessible() Then
        GetCodeModuleCount = Application.VBE.ActiveVBProject.VBComponents.Count
    Else
        GetCodeModuleCount = 0
    End If
End Function

Private Sub GetTableCounts(ByRef linkedCount As Long, ByRef localCount As Long)
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 1) <> "~" Then
            If Len(tdf.Connect & "") > 0 Then
                linkedCount = linkedCount + 1
            Else
                localCount = localCount + 1
            End If
        End If
    Next
End Sub

Private Function GetBackendList() As String
    Dim tdf As DAO.TableDef
    Dim dict As Object, key As Variant
    Dim conn As String, s As String

    Set dict = CreateObject("Scripting.Dictionary")

    For Each tdf In CurrentDb.TableDefs
        If Len(tdf.Connect & "") > 0 Then
            conn = tdf.Connect
            If Not dict.Exists(conn) Then dict.Add conn, 1 Else dict(conn) = dict(conn) + 1
        End If
    Next

    For Each key In dict.Keys
        s = s & "  - " & SummarizeConnection(CStr(key)) & " (" & dict(key) & " tables)" & vbCrLf
    Next

    GetBackendList = s
End Function

Private Function SummarizeConnection(ByVal conn As String) As String
    Dim u As String: u = UCase$(conn)
    If InStr(u, "ODBC;") > 0 Then
        SummarizeConnection = "ODBC: " & ExtractConnValue(conn, "DATABASE", ExtractConnValue(conn, "DSN", conn))
    ElseIf InStr(u, "DATABASE=") > 0 Then
        SummarizeConnection = ExtractConnValue(conn, "DATABASE", conn)
    Else
        SummarizeConnection = conn
    End If
End Function

Private Function ExtractConnValue(ByVal conn As String, ByVal key As String, ByVal fallback As String) As String
    Dim parts() As String, i As Long, kv() As String
    parts = Split(conn, ";")
    For i = LBound(parts) To UBound(parts)
        If UCase$(Left$(Trim$(parts(i)), Len(key) + 1)) = UCase$(key & "=") Then
            kv = Split(parts(i), "=")
            If UBound(kv) >= 1 Then
                ExtractConnValue = Trim$(kv(1))
                Exit Function
            End If
        End If
    Next
    ExtractConnValue = fallback
End Function

Private Function GetReferences() As String
    On Error GoTo ErrHandler
    Dim ref As Reference, s As String
    For Each ref In Application.References
        s = s & "  " & IIf(ref.IsBroken, "[BROKEN] ", "") & ref.Name
        On Error Resume Next
        s = s & " - " & ref.FullPath
        On Error GoTo ErrHandler
        s = s & vbCrLf
    Next
    GetReferences = s
    Exit Function
ErrHandler:
    GetReferences = "  (Unable to enumerate references)" & vbCrLf
End Function

Private Function GetFieldTypeName(ByVal t As Integer) As String
    Select Case t
        Case dbBoolean: GetFieldTypeName = "Yes/No"
        Case dbByte: GetFieldTypeName = "Byte"
        Case dbInteger: GetFieldTypeName = "Integer"
        Case dbLong: GetFieldTypeName = "Long"
        Case dbCurrency: GetFieldTypeName = "Currency"
        Case dbSingle: GetFieldTypeName = "Single"
        Case dbDouble: GetFieldTypeName = "Double"
        Case dbDate: GetFieldTypeName = "Date/Time"
        Case dbText: GetFieldTypeName = "Text"
        Case dbLongBinary: GetFieldTypeName = "OLE Object"
        Case dbMemo: GetFieldTypeName = "Memo"
        Case dbGUID: GetFieldTypeName = "GUID"
        Case dbBigInt: GetFieldTypeName = "BigInt"
        Case dbVarBinary: GetFieldTypeName = "VarBinary"
        Case dbChar: GetFieldTypeName = "Char"
        Case dbNumeric: GetFieldTypeName = "Numeric"
        Case dbDecimal: GetFieldTypeName = "Decimal"
        Case dbFloat: GetFieldTypeName = "Float"
        Case dbTime: GetFieldTypeName = "Time"
        Case dbTimeStamp: GetFieldTypeName = "TimeStamp"
        Case dbAttachment: GetFieldTypeName = "Attachment"
        Case dbComplexByte, dbComplexInteger, dbComplexLong, dbComplexSingle, _
             dbComplexDouble, dbComplexGUID, dbComplexDecimal, dbComplexText
            GetFieldTypeName = "MultiValue"
        Case Else: GetFieldTypeName = "Unknown(" & t & ")"
    End Select
End Function

Private Function GetQueryTypeName(ByVal t As Integer) As String
    Select Case t
        Case 0: GetQueryTypeName = "Select"
        Case 16: GetQueryTypeName = "Crosstab"
        Case 32: GetQueryTypeName = "Delete"
        Case 48: GetQueryTypeName = "Update"
        Case 64: GetQueryTypeName = "Append"
        Case 80: GetQueryTypeName = "Make-Table"
        Case 96: GetQueryTypeName = "Data-Definition"
        Case 112: GetQueryTypeName = "Pass-Through"
        Case 128: GetQueryTypeName = "Union"
        Case Else: GetQueryTypeName = "Unknown(" & t & ")"
    End Select
End Function

Private Function GetIndexFields(ByVal idx As DAO.Index) As String
    Dim fld As DAO.Field, s As String
    For Each fld In idx.Fields
        s = s & fld.Name & ", "
    Next
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)
    GetIndexFields = s
End Function

Private Function GetSaveAsTextContent(ByVal objType As AcObjectType, ByVal objName As String) As String
    On Error GoTo ErrHandler
    Dim tmp As String, content As String

    tmp = GetTempFolder() & "sav_" & objName & "_" & Format(Now, "yymmddhhnnss") & ".txt"

    Application.SaveAsText objType, objName, tmp
    content = ReadFile(tmp)

    On Error Resume Next
    Kill tmp
    On Error GoTo 0

    GetSaveAsTextContent = content
    Exit Function

ErrHandler:
    GetSaveAsTextContent = "(SaveAsText failed: " & Err.Number & " - " & Err.Description & ")"
End Function

Private Function IsVBEAccessible() As Boolean
    On Error Resume Next
    Dim dummy As Long
    dummy = Application.VBE.ActiveVBProject.VBComponents.Count
    IsVBEAccessible = (Err.Number = 0)
End Function

Private Function GetTempFolder() As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetTempFolder = fso.GetSpecialFolder(2).Path & "\"
    If Len(GetTempFolder) = 0 Then GetTempFolder = Environ$("TEMP") & "\"
    If Len(GetTempFolder) = 1 Then GetTempFolder = CurrentProject.Path & "\"
End Function

Private Function BuildPath(ByVal folder As String, ByVal fileName As String) As String
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    BuildPath = folder & fileName
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Function ReadFile(ByVal path As String) As String
    Dim f As Integer, line As String, content As String
    f = FreeFile
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        content = content & line & vbCrLf
    Loop
    Close #f
    ReadFile = content
End Function

Private Sub WriteFile(ByVal path As String, ByVal content As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, content
    Close #f
End Sub

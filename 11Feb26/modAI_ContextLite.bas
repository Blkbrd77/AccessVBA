Attribute VB_Name = "modAI_ContextLite"
Option Compare Database
Option Explicit

' ============================================================
' AI Context Lite - Streamlined export for AI assistants
' Produces ~1000-1500 lines vs 8000+ from full export
'
' Entry point:
'   Call ExportContextLite
'   Call ExportContextLite "C:\MyFolder"
' ============================================================

Public Sub ExportContextLite(Optional ByVal OutputFolder As String = "")
    On Error GoTo ErrHandler

    If Len(OutputFolder) = 0 Then OutputFolder = CurrentProject.Path

    Dim report As String, outFile As String
    outFile = BuildPath(OutputFolder, "ContextLite_" & Format(Now, "yyyymmdd") & ".txt")

    report = BuildHeader()
    report = report & BuildTableSection()
    report = report & BuildRelationshipsSection()
    report = report & BuildQueriesSection()
    report = report & BuildFormsLite()
    report = report & BuildVBASignatures()
    report = report & BuildDeprecationHints()

    WriteFile outFile, report
    MsgBox "Exported to:" & vbCrLf & outFile, vbInformation, "Context Lite"
    Exit Sub

ErrHandler:
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

' ============================================================
' Header - Quick stats
' ============================================================

Private Function BuildHeader() As String
    Dim s As String
    s = "# AI CONTEXT LITE - " & CurrentProject.Name & vbCrLf
    s = s & "Generated: " & Now & vbCrLf
    s = s & "Tables: " & GetUserTableCount() & " | "
    s = s & "Queries: " & CurrentData.AllQueries.Count & " | "
    s = s & "Forms: " & CurrentProject.AllForms.Count & " | "
    s = s & "Modules: " & GetModuleCount() & vbCrLf
    s = s & String(60, "-") & vbCrLf & vbCrLf
    BuildHeader = s
End Function

' ============================================================
' Tables - Schema with row counts, compact format
' ============================================================

Private Function BuildTableSection() As String
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tdf As DAO.TableDef, fld As DAO.Field, idx As DAO.Index
    Dim s As String, flds As String, pk As String, rowCount As Long

    s = "## TABLES" & vbCrLf & vbCrLf

    For Each tdf In db.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 1) <> "~" Then
            ' Row count
            On Error Resume Next
            rowCount = DCount("*", tdf.Name)
            If Err.Number <> 0 Then rowCount = -1
            On Error GoTo 0

            s = s & "### " & tdf.Name & " (" & rowCount & " rows)"
            If Len(tdf.Connect) > 0 Then s = s & " [LINKED]"
            s = s & vbCrLf

            ' Fields - compact list
            flds = ""
            For Each fld In tdf.Fields
                flds = flds & "  " & fld.Name & ": " & GetTypeName(fld.Type)
                If fld.Type = dbText And fld.Size < 255 Then flds = flds & "(" & fld.Size & ")"
                If fld.Required Then flds = flds & " NOT NULL"
                If Len(fld.DefaultValue) > 0 Then flds = flds & " DEFAULT " & fld.DefaultValue
                flds = flds & vbCrLf
            Next
            s = s & flds

            ' Primary key only
            pk = ""
            For Each idx In tdf.Indexes
                If idx.Primary Then
                    pk = GetIdxFields(idx)
                    Exit For
                End If
            Next
            If Len(pk) > 0 Then s = s & "  PK: " & pk & vbCrLf

            s = s & vbCrLf
        End If
    Next

    BuildTableSection = s
End Function

' ============================================================
' Relationships - Compact format
' ============================================================

Private Function BuildRelationshipsSection() As String
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rel As DAO.Relation, fld As DAO.Field
    Dim s As String, attrs As String

    s = "## RELATIONSHIPS" & vbCrLf & vbCrLf

    If db.Relations.Count = 0 Then
        s = s & "(none)" & vbCrLf & vbCrLf
        BuildRelationshipsSection = s
        Exit Function
    End If

    For Each rel In db.Relations
        s = s & rel.Table & "." & rel.Fields(0).Name & " -> "
        s = s & rel.ForeignTable & "." & rel.Fields(0).ForeignName

        ' Attributes inline
        If (rel.Attributes And dbRelationDontEnforce) = 0 Then
            s = s & " [Enforced"
            If (rel.Attributes And dbRelationUpdateCascade) <> 0 Then s = s & ", CascadeUpdate"
            If (rel.Attributes And dbRelationDeleteCascade) <> 0 Then s = s & ", CascadeDelete"
            s = s & "]"
        End If
        s = s & vbCrLf
    Next

    s = s & vbCrLf
    BuildRelationshipsSection = s
End Function

' ============================================================
' Queries - SQL only
' ============================================================

Private Function BuildQueriesSection() As String
    Dim db As DAO.Database: Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim s As String

    s = "## QUERIES" & vbCrLf & vbCrLf

    For Each qdf In db.QueryDefs
        If Left$(qdf.Name, 1) <> "~" Then
            s = s & "### " & qdf.Name & vbCrLf
            s = s & "```sql" & vbCrLf
            s = s & Trim$(qdf.SQL) & vbCrLf
            s = s & "```" & vbCrLf & vbCrLf
        End If
    Next

    BuildQueriesSection = s
End Function

' ============================================================
' Forms Lite - Just RecordSource and control bindings
' ============================================================

Private Function BuildFormsLite() As String
    Dim frm As AccessObject
    Dim s As String

    s = "## FORMS" & vbCrLf & vbCrLf

    For Each frm In CurrentProject.AllForms
        s = s & GetFormSummary(frm.Name)
    Next

    BuildFormsLite = s
End Function

Private Function GetFormSummary(ByVal frmName As String) As String
    On Error GoTo ErrHandler

    Dim s As String, ctl As Control
    Dim wasOpen As Boolean, src As String

    wasOpen = (SysCmd(acSysCmdGetObjectState, acForm, frmName) <> 0)
    If Not wasOpen Then DoCmd.OpenForm frmName, acDesign, , , , acHidden

    s = "### " & frmName & vbCrLf

    ' RecordSource
    src = Nz(Forms(frmName).RecordSource, "(unbound)")
    s = s & "RecordSource: " & src & vbCrLf

    ' Bound controls only
    s = s & "Controls:" & vbCrLf
    For Each ctl In Forms(frmName).Controls
        On Error Resume Next
        If Len(ctl.ControlSource) > 0 Then
            s = s & "  " & ctl.Name & " <- " & ctl.ControlSource & vbCrLf
        End If
        On Error GoTo ErrHandler
    Next

    ' Check for code behind
    If HasFormCode(frmName) Then
        s = s & "HasCode: Yes" & vbCrLf
    End If

    If Not wasOpen Then DoCmd.Close acForm, frmName, acSaveNo

    s = s & vbCrLf
    GetFormSummary = s
    Exit Function

ErrHandler:
    If Not wasOpen Then
        On Error Resume Next
        DoCmd.Close acForm, frmName, acSaveNo
    End If
    GetFormSummary = "### " & frmName & " (error reading)" & vbCrLf & vbCrLf
End Function

Private Function HasFormCode(ByVal frmName As String) As Boolean
    On Error Resume Next
    Dim lineCount As Long
    lineCount = Application.VBE.ActiveVBProject.VBComponents("Form_" & frmName).CodeModule.CountOfLines
    HasFormCode = (lineCount > 2)  ' More than just Option Compare/Explicit
End Function

' ============================================================
' VBA Signatures - Public Subs/Functions only
' ============================================================

Private Function BuildVBASignatures() As String
    Dim s As String

    s = "## VBA PUBLIC INTERFACE" & vbCrLf & vbCrLf

    If Not IsVBEAccessible() Then
        s = s & "(VBA project access disabled - enable in Trust Center)" & vbCrLf & vbCrLf
        BuildVBASignatures = s
        Exit Function
    End If

    On Error GoTo ErrHandler

    Dim vbComp As Object, codeMod As Object
    Dim i As Long, lineText As String, inProc As Boolean
    Dim modType As String

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        ' Skip form/report modules - already covered
        If vbComp.Type = 1 Then  ' Standard module only
            Set codeMod = vbComp.CodeModule

            If codeMod.CountOfLines > 0 Then
                s = s & "### " & vbComp.Name & vbCrLf

                For i = 1 To codeMod.CountOfLines
                    lineText = Trim$(codeMod.Lines(i, 1))

                    ' Public Sub/Function declarations
                    If StartsWith(lineText, "Public Sub ") Or _
                       StartsWith(lineText, "Public Function ") Or _
                       StartsWith(lineText, "Sub ") Or _
                       StartsWith(lineText, "Function ") Then

                        ' Get full signature (may span lines with _)
                        Dim sig As String: sig = lineText
                        Do While Right$(Trim$(sig), 1) = "_" And i < codeMod.CountOfLines
                            i = i + 1
                            sig = Left$(sig, Len(sig) - 1) & " " & Trim$(codeMod.Lines(i, 1))
                        Loop

                        s = s & "  " & sig & vbCrLf
                    End If
                Next

                s = s & vbCrLf
            End If
        End If
    Next

    BuildVBASignatures = s
    Exit Function

ErrHandler:
    s = s & "(Error reading VBA: " & Err.Description & ")" & vbCrLf & vbCrLf
    BuildVBASignatures = s
End Function

' ============================================================
' Deprecation Hints - Empty tables, orphaned objects
' ============================================================

Private Function BuildDeprecationHints() As String
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tdf As DAO.TableDef
    Dim s As String, rowCount As Long
    Dim emptyTables As String, unusedCount As Long

    s = "## DEPRECATION HINTS" & vbCrLf & vbCrLf

    ' Empty tables
    s = s & "### Empty Tables (0 rows)" & vbCrLf
    For Each tdf In db.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 1) <> "~" Then
            On Error Resume Next
            rowCount = DCount("*", tdf.Name)
            If Err.Number = 0 And rowCount = 0 Then
                s = s & "  - " & tdf.Name & vbCrLf
                unusedCount = unusedCount + 1
            End If
            On Error GoTo 0
        End If
    Next
    If unusedCount = 0 Then s = s & "  (none)" & vbCrLf

    s = s & vbCrLf

    ' Tables with tmp/Import prefix (likely temporary)
    s = s & "### Likely Temporary Tables" & vbCrLf
    unusedCount = 0
    For Each tdf In db.TableDefs
        If Left$(tdf.Name, 3) = "tmp" Or Left$(tdf.Name, 3) = "tbl" & "Import" Or _
           Left$(tdf.Name, 6) = "Import" Then
            s = s & "  - " & tdf.Name & vbCrLf
            unusedCount = unusedCount + 1
        End If
    Next
    If unusedCount = 0 Then s = s & "  (none)" & vbCrLf

    s = s & vbCrLf

    ' Tables not referenced in any query
    s = s & "### Tables Not in Queries" & vbCrLf
    Dim allSQL As String, qdf As DAO.QueryDef
    For Each qdf In db.QueryDefs
        allSQL = allSQL & " " & UCase$(qdf.SQL) & " "
    Next

    unusedCount = 0
    For Each tdf In db.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 1) <> "~" Then
            If InStr(allSQL, UCase$(tdf.Name)) = 0 Then
                s = s & "  - " & tdf.Name & vbCrLf
                unusedCount = unusedCount + 1
            End If
        End If
    Next
    If unusedCount = 0 Then s = s & "  (none)" & vbCrLf

    s = s & vbCrLf
    BuildDeprecationHints = s
End Function

' ============================================================
' Utilities
' ============================================================

Private Function GetUserTableCount() As Long
    Dim tdf As DAO.TableDef, cnt As Long
    For Each tdf In CurrentDb.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 1) <> "~" Then cnt = cnt + 1
    Next
    GetUserTableCount = cnt
End Function

Private Function GetModuleCount() As Long
    On Error Resume Next
    If IsVBEAccessible() Then
        GetModuleCount = Application.VBE.ActiveVBProject.VBComponents.Count
    End If
End Function

Private Function GetTypeName(ByVal t As Integer) As String
    Select Case t
        Case dbBoolean: GetTypeName = "Yes/No"
        Case dbByte: GetTypeName = "Byte"
        Case dbInteger: GetTypeName = "Integer"
        Case dbLong: GetTypeName = "Long"
        Case dbCurrency: GetTypeName = "Currency"
        Case dbSingle: GetTypeName = "Single"
        Case dbDouble: GetTypeName = "Double"
        Case dbDate: GetTypeName = "Date/Time"
        Case dbText: GetTypeName = "Text"
        Case dbMemo: GetTypeName = "Memo"
        Case dbGUID: GetTypeName = "GUID"
        Case Else: GetTypeName = "Type" & t
    End Select
End Function

Private Function GetIdxFields(ByVal idx As DAO.Index) As String
    Dim fld As DAO.Field, s As String
    For Each fld In idx.Fields
        s = s & fld.Name & ", "
    Next
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)
    GetIdxFields = s
End Function

Private Function IsVBEAccessible() As Boolean
    On Error Resume Next
    Dim dummy As Long
    dummy = Application.VBE.ActiveVBProject.VBComponents.Count
    IsVBEAccessible = (Err.Number = 0)
End Function

Private Function StartsWith(ByVal str As String, ByVal prefix As String) As Boolean
    StartsWith = (Left$(str, Len(prefix)) = prefix)
End Function

Private Function BuildPath(ByVal folder As String, ByVal fileName As String) As String
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    BuildPath = folder & fileName
End Function

Private Sub WriteFile(ByVal path As String, ByVal content As String)
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, content
    Close #f
End Sub

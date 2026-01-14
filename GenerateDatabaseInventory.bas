
### 2. VBA Inventory Script  
**File name suggestion:** `access-inventory-export.bas` or `GenerateDatabaseInventory.bas`  
This is the VBA code we reviewed earlier, cleaned up slightly for GitHub readability. Save it with a `.bas` extension so Access can import it easily if needed.

```vb
Option Compare Database
Option Explicit

' VBA script to generate a plain-text inventory of Access database objects
' Focus: Tables (with fields), Queries (with SQL), Forms, Reports, Modules
' Output: Text file with timestamp in filename
' Usage: Run GenerateDatabaseInventory from VBA editor or macro

Public Sub GenerateDatabaseInventory()
    On Error GoTo ErrorHandler
    
    Dim fso         As Object
    Dim txtFile     As Object
    Dim strFilePath As String
    Dim db          As DAO.Database
    Dim tdf         As DAO.TableDef
    Dim qdf         As DAO.QueryDef
    Dim fld         As DAO.Field
    Dim strLine     As String
    
    ' === CHANGE THIS PATH to a folder you can write to ===
    strFilePath = "C:\Temp\AccessInventory_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(strFilePath, True)
    Set db = CurrentDb
    
    txtFile.WriteLine "Database Inventory Report"
    txtFile.WriteLine "Generated: " & Now
    txtFile.WriteLine "Database: " & CurrentProject.FullName
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine vbCrLf
    
    ' TABLES
    txtFile.WriteLine "TABLES"
    txtFile.WriteLine String(40, "-")
    
    For Each tdf In db.TableDefs
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            txtFile.WriteLine vbCrLf & "Table: " & tdf.Name
            
            If Len(tdf.Connect) > 0 Then
                txtFile.WriteLine "  Type: Linked (" & tdf.Connect & ")"
            Else
                txtFile.WriteLine "  Type: Local"
            End If
            
            txtFile.WriteLine "  Fields:"
            For Each fld In tdf.Fields
                strLine = "    - " & fld.Name & vbTab & fld.Type & " (" & fld.Size & ")"
                If fld.Required Then strLine = strLine & " [Required]"
                If fld.DefaultValue <> "" Then strLine = strLine & " [Default: " & fld.DefaultValue & "]"
                txtFile.WriteLine strLine
            Next fld
        End If
    Next tdf
    
    txtFile.WriteLine vbCrLf & String(80, "=") & vbCrLf
    
    ' QUERIES
    txtFile.WriteLine "QUERIES"
    txtFile.WriteLine String(40, "-")
    
    For Each qdf In db.QueryDefs
        If Left(qdf.Name, 1) <> "~" And Left(qdf.Name, 4) <> "MSys" Then
            txtFile.WriteLine vbCrLf & "Query: " & qdf.Name
            txtFile.WriteLine "  SQL:"
            txtFile.WriteLine "  " & Replace(qdf.SQL, vbCrLf, vbCrLf & "  ")
        End If
    Next qdf
    
    txtFile.WriteLine vbCrLf & String(80, "=") & vbCrLf
    
    ' FORMS, REPORTS, MODULES (basic list)
    txtFile.WriteLine "FORMS"
    txtFile.WriteLine String(40, "-")
    Dim accObj As AccessObject
    For Each accObj In CurrentProject.AllForms
        txtFile.WriteLine "Form: " & accObj.Name & "  (Modified: " & accObj.DateModified & ")"
    Next accObj
    
    txtFile.WriteLine vbCrLf & "REPORTS"
    txtFile.WriteLine String(40, "-")
    For Each accObj In CurrentProject.AllReports
        txtFile.WriteLine "Report: " & accObj.Name & "  (Modified: " & accObj.DateModified & ")"
    Next accObj
    
    txtFile.WriteLine vbCrLf & "MODULES"
    txtFile.WriteLine String(40, "-")
    For Each accObj In CurrentProject.AllModules
        txtFile.WriteLine "Module: " & accObj.Name & "  (Modified: " & accObj.DateModified & ")"
    Next accObj
    
    txtFile.WriteLine vbCrLf & "End of Report"
    
Cleanup:
    txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing
    Set db = Nothing
    MsgBox "Inventory saved to:" & vbCrLf & strFilePath, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Cleanup
End Sub

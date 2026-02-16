Attribute VB_Name = "basOrderNumbering"
Option Compare Database
Option Explicit

'==== Public API ====
Public Function NextSalesBaseToken() As Long
    ' Returns the next BaseToken for a SALES order for the current year
    On Error GoTo EH

    Dim yr As Integer
    yr = Year(Date)  ' use today's date for the annual seed

    Dim seed As Long, upper As Long
    seed = SalesBaseSeedForYear(yr)
    upper = seed + 999

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim maxVal As Long

    Set db = CurrentDb

    ' Use VAL(BaseToken) so it works whether BaseToken is Text or Number
    sql = "SELECT MAX(VAL(BaseToken)) AS MaxBT " & _
          "FROM SalesOrders " & _
          "WHERE OrderType='SALES' " & _
          "AND VAL(BaseToken) BETWEEN " & seed & " AND " & upper & ";"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        maxVal = Nz(rs!MaxBT, 0)
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    If maxVal = 0 Then
        NextSalesBaseToken = seed          ' no orders yet this year
    ElseIf maxVal < upper Then
        NextSalesBaseToken = maxVal + 1    ' next within this year's block
    Else
        Err.Raise vbObjectError + 513, "NextSalesBaseToken", _
                  "Yearly SALES BaseToken block exhausted (" & yr & ")."
    End If

    Exit Function

EH:
    MsgBox "NextSalesBaseToken error: " & Err.Description, vbExclamation
    NextSalesBaseToken = 0
End Function

Public Function DefaultSalesSystemLetter() As String
    ' Default for SALES orders, user may override on form if needed
    DefaultSalesSystemLetter = "P"
End Function

'==== Internal helpers ====
Private Function SalesBaseSeedForYear(ByVal yr As Integer) As Long
    ' 2026 -> 576080, 2027 -> 577080, i.e. +1000 per year
    ' => seed = 576080 + (yr - 2026) * 1000
    SalesBaseSeedForYear = 576080 + CLng(yr - 2026) * 1000
End Function

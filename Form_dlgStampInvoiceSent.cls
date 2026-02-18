VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dlgStampInvoiceSent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=== Form: dlgStampInvoiceSent ===
' Purpose : Modal dialog to capture the Proforma Invoice Date for an order.
'           Sets TempVars("InvoiceSentDate") and TempVars("InvoiceSentResult").
'
' Caller  : frmOrderList.cmdProformaInvoiceDate_Click()
'
' Controls required on the form:
'   txtStampInvoiceDateInput  - Text box (bound to nothing, format: Short Date)
'   cmdOK                     - Command button
'   cmdCancel                 - Command button
'
Option Compare Database
Option Explicit

Private Sub cmdOK_Click()
    On Error GoTo EH

    Dim v As Variant
    v = Nz(Me!txtStampInvoiceDateInput.Value, Null)

    If IsNull(v) Or Len(Trim$(v & "")) = 0 Then
        MsgBox "Please enter the proforma invoice date.", vbExclamation, "Missing Date"
        Me!txtStampInvoiceDateInput.SetFocus
        Exit Sub
    End If

    If Not IsDate(v) Then
        MsgBox "That doesn't look like a valid date. Try something like 1/23/2026.", _
               vbExclamation, "Invalid Date"
        Me!txtStampInvoiceDateInput.SetFocus
        Exit Sub
    End If

    ' Store as a pure Date (no time component)
    TempVars("InvoiceSentDate") = DateValue(CDate(v))

    ' Signal success to the caller
    TempVars("InvoiceSentResult") = "OK"

    DoCmd.Close acForm, Me.Name
    Exit Sub

EH:
    MsgBox "Error stamping invoice date: " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancel_Click()
    TempVars("InvoiceSentResult") = "Cancel"
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
    ' Default to today so the user only has to change it if needed
    Me!txtStampInvoiceDateInput = Date

    ' Clear any leftover TempVars from a previous run
    On Error Resume Next
    TempVars.Remove "InvoiceSentDate"
    TempVars.Remove "InvoiceSentResult"
    On Error GoTo 0
End Sub

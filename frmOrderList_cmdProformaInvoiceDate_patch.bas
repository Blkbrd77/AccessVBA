'==============================================================================
' PATCH: frmOrderList - Proforma Invoice Date button handler
'
' HOW TO APPLY
' ------------
' 1. Open frmOrderList in the Access VBA editor (Alt+F11 -> select the form).
' 2. Add a new Command Button to the form.
'      Name    : cmdProformaInvoiceDate
'      Caption : Proforma Invoice Date
' 3. Add a new unbound Text Box to the form to display the stored date.
'      Name           : txtProformaInvoiceDate
'      Control Source : DateInvoiceSent      (bound to the field)
'      Format         : Short Date           (or your preferred date format)
'      Locked         : Yes / Enabled: No    (display-only; editing done via dialog)
' 4. Paste the Sub below into the form's code module.
' 5. Also add the one-line stub in Form_Current (shown at the bottom).
'==============================================================================

Option Compare Database
Option Explicit

' ---------------------------------------------------------------------------
' cmdProformaInvoiceDate_Click
' Opens the dlgStampInvoiceSent dialog, captures the chosen date, writes it
' to DateInvoiceSent, and logs the change to tblOrderAudit.
' ---------------------------------------------------------------------------
Private Sub cmdProformaInvoiceDate_Click()
    On Error GoTo EH

    Dim oldDateInvoiceSent As Variant
    Dim newDateInvoiceSent As Variant
    Dim lngSOID As Long
    Dim sOrderNumber As String

    Debug.Print "=== STAMP INVOICE SENT START ==="

    ' Must be on an existing, saved record
    If Me.NewRecord Then
        MsgBox "Please select or save a record before stamping.", vbExclamation
        Exit Sub
    End If

    ' Capture current state BEFORE making any changes
    lngSOID = Nz(Me!SOID, 0)
    sOrderNumber = Nz(Me!OrderNumber, "")
    oldDateInvoiceSent = Me!DateInvoiceSent  ' May be Null

    Debug.Print "SOID: " & lngSOID
    Debug.Print "OrderNumber: " & sOrderNumber
    Debug.Print "Old DateInvoiceSent: " & IIf(IsNull(oldDateInvoiceSent), "(null)", oldDateInvoiceSent)

    ' Clear any leftover TempVars
    On Error Resume Next
    TempVars.Remove "InvoiceSentDate"
    TempVars.Remove "InvoiceSentResult"
    On Error GoTo EH

    ' Open the dialog modally - execution resumes here after it closes
    DoCmd.OpenForm "dlgStampInvoiceSent", WindowMode:=acDialog

    ' Check whether the user confirmed or cancelled
    If Nz(TempVars("InvoiceSentResult"), "Cancel") <> "OK" Then
        Debug.Print "User canceled invoice sent dialog"
        Exit Sub
    End If

    ' Retrieve the chosen date from TempVars
    newDateInvoiceSent = TempVars("InvoiceSentDate")
    Debug.Print "New DateInvoiceSent: " & newDateInvoiceSent

    ' Write to the bound field
    Me!DateInvoiceSent = newDateInvoiceSent

    ' Persist immediately
    If Me.Dirty Then Me.Dirty = False

    Debug.Print "About to call LogOrderAction..."

    ' ---- AUDIT: Stamp success ----
    Call LogOrderAction(lngSOID, sOrderNumber, "STAMP_INVOICE_SENT", _
                   IIf(IsNull(oldDateInvoiceSent), "", Format(oldDateInvoiceSent, "yyyy-mm-dd")), _
                   Format(newDateInvoiceSent, "yyyy-mm-dd"), _
                   "Proforma invoice date stamped by user")

    Debug.Print "LogOrderAction called successfully"

    ' Refresh the display text box
    Me!txtProformaInvoiceDate.Requery

    Debug.Print "=== STAMP INVOICE SENT COMPLETE ==="
    Exit Sub

EH:
    Debug.Print "=== ERROR IN STAMP INVOICE SENT ==="
    Debug.Print "Error: " & Err.Number & " - " & Err.Description

    ' ---- AUDIT: Stamp failure ----
    On Error Resume Next
    Call LogOrderAction(Nz(Me!SOID, 0), Nz(Me!OrderNumber, ""), "STAMP_INVOICE_SENT_FAILED", "", "", _
                   "Err " & Err.Number & ": " & Err.Description)
    On Error GoTo 0

    MsgBox "Stamp Invoice Sent failed: " & Err.Description, vbExclamation
End Sub


' ---------------------------------------------------------------------------
' ADD TO Form_Current (inside the existing Sub, alongside the other button
' enable/disable logic)
' ---------------------------------------------------------------------------
'    Me.cmdProformaInvoiceDate.Enabled = (Not IsNull(Me!SOID))

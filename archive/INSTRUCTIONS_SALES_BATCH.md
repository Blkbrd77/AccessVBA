# Instructions: Enable SALES Orders in Batch Generator

Based on review of `numberResContext` export from `Q1019_23Jan26_Start_travellaptop.accdb`.

## Overview

Currently the batch generation flow supports PROJECT orders with qualifiers. To support SALES orders with the unified format (`<BaseToken>-<QualifierCode><SystemLetter><Seq>-<BackorderNo>`), you need to:

1. Update the form UI to handle SALES-specific BaseToken generation
2. Ensure the backend modules already support SALES (most do via the `scope` parameter)
3. Add logic to auto-generate SALES BaseTokens using the year-based scheme

---

## Step 1: Modify `dlgBatchGenerateOrders` Form (or `frmSalesOrderEntry`)

**Add/Modify Controls:**

```
┌─────────────────────────────────────────────────────────────┐
│  Order Type:    [cboOrderType ▼]  ◄── "SALES" or "PROJECT"  │
│                                                             │
│  ── If SALES ──────────────────────────────────────────     │
│  Base Token:    [txtBaseToken]  [btnGetNextBase]            │
│                  └── Auto-fills with next available         │
│                      (e.g., 576001)                         │
│                                                             │
│  ── If PROJECT ────────────────────────────────────────     │
│  Project:       [cboProjectSelect ▼]                        │
│                  └── Populates txtBaseToken from Projects   │
│                      table (e.g., TAGO25)                   │
└─────────────────────────────────────────────────────────────┘
```

**Form Code Changes (`Form_dlgBatchGenerateOrders` or `Form_frmSalesOrderEntry`):**

```vba
'===============================================
' cboOrderType_AfterUpdate
' Show/hide appropriate controls based on type
'===============================================
Private Sub cboOrderType_AfterUpdate()
    Dim isSales As Boolean
    isSales = (Nz(Me.cboOrderType, "") = "SALES")

    ' Toggle visibility
    Me.cboProjectSelect.Visible = Not isSales
    Me.lblProject.Visible = Not isSales
    Me.btnGetNextBase.Visible = isSales

    ' Clear and set BaseToken
    If isSales Then
        Me.txtBaseToken = ""
        Me.txtBaseToken.Locked = False
    Else
        Me.txtBaseToken = ""
        Me.txtBaseToken.Locked = True  ' Filled by project selection
    End If
End Sub

'===============================================
' btnGetNextBase_Click
' Auto-generate next SALES BaseToken
'===============================================
Private Sub btnGetNextBase_Click()
    On Error GoTo EH

    If Nz(Me.cboOrderType, "") <> "SALES" Then
        MsgBox "This button is only for SALES orders.", vbExclamation
        Exit Sub
    End If

    Dim nextBase As Long
    nextBase = NextSalesBaseToken()  ' From basOrderNumbering

    If nextBase = 0 Then
        MsgBox "Could not generate BaseToken.", vbExclamation
        Exit Sub
    End If

    Me.txtBaseToken = CStr(nextBase)
    Exit Sub

EH:
    MsgBox "Error getting next BaseToken: " & Err.Description, vbExclamation
End Sub

'===============================================
' cboProjectSelect_AfterUpdate
' Populate BaseToken from selected project
'===============================================
Private Sub cboProjectSelect_AfterUpdate()
    If Nz(Me.cboOrderType, "") = "PROJECT" Then
        Me.txtBaseToken = Nz(Me.cboProjectSelect.Column(0), "")  ' BaseToken column
    End If
End Sub
```

---

## Step 2: Update Form_Open to Initialize Correctly

```vba
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo EH

    ' Set default order type
    If IsNull(Me.cboOrderType) Then Me.cboOrderType = "PROJECT"

    ' Trigger visibility logic
    Call cboOrderType_AfterUpdate

    ' Reset qualifier quantities
    Reset_Qty_From_Qualifiers
    Me.subQty.Requery

    ' Clear preview
    CurrentDb.Execute "DELETE FROM tmpBatchPreview;", dbFailOnError
    Me.lstPreview.Requery

    Exit Sub
EH:
    MsgBox "Open failed: " & Err.Description, vbExclamation
End Sub
```

---

## Step 3: Update Validation in `ValidateHeader`

```vba
Private Function ValidateHeader(ByRef scope As String, _
                                 ByRef baseT As String, _
                                 ByRef sysLet As String) As Boolean
    ValidateHeader = False

    scope = Nz(Me.cboOrderType, "")
    baseT = Nz(Me.txtBaseToken, "")
    sysLet = Nz(Me.cboSystemLetter, "")

    If Len(scope) = 0 Then
        MsgBox "Order Type is required.", vbExclamation
        Exit Function
    End If

    If Len(baseT) = 0 Then
        MsgBox "Base Token is required." & vbCrLf & _
               IIf(scope = "SALES", "Click 'Get Next' to generate.", "Select a Project."), _
               vbExclamation
        Exit Function
    End If

    If Len(sysLet) = 0 Then
        MsgBox "System Letter is required.", vbExclamation
        Exit Function
    End If

    ' SALES-specific validation: BaseToken should be numeric
    If scope = "SALES" Then
        If Not IsNumeric(baseT) Then
            MsgBox "SALES BaseToken must be numeric (e.g., 576001).", vbExclamation
            Exit Function
        End If
    End If

    ValidateHeader = True
End Function
```

---

## Step 4: Verify Backend Modules (No Changes Needed)

These modules already use a `scope` parameter and should work for both SALES and PROJECT:

| Module | Function | Status |
|--------|----------|--------|
| `basSeqAllocator` | `ReserveSeq(scope, BaseToken, QualifierCode, SystemLetter)` | Ready |
| `basBatchWizard` | `BuildPreview(scope, BaseToken, SystemLetter)` | Ready |
| `basBatchWizard` | `CommitBatch(scope, BaseToken, SystemLetter, ...)` | Ready |
| `basBatchWizard` | `CreateOneOrder(scope, BaseToken, QualifierCode, ...)` | Ready |

---

## Step 5: Add "Get Next Base" Button to Form Design

In the Access form designer:

1. Open `dlgBatchGenerateOrders` (or `frmSalesOrderEntry`) in Design View
2. Add a **Command Button** next to `txtBaseToken`:
   - **Name**: `btnGetNextBase`
   - **Caption**: `Get Next` or `Auto`
   - **On Click**: `[Event Procedure]` → add `btnGetNextBase_Click` code above
3. Add a **Label** for the Project combo if not present:
   - **Name**: `lblProject`
   - **Caption**: `Project:`
4. Save the form

---

## Step 6: Update `cboOrderType` Row Source

Ensure the Order Type combo includes both options:

```
Row Source Type: Value List
Row Source: "SALES";"PROJECT"
```

---

## Step 7: Test the Flow

**Test SALES:**
1. Open `dlgBatchGenerateOrders`
2. Select Order Type = "SALES"
3. Click "Get Next" → should populate BaseToken (e.g., `576001`)
4. Select System Letter (e.g., `N`)
5. Enter quantities for qualifiers (e.g., CM=3, CE=2)
6. Click "Preview Numbers" → should show:
   ```
   576001-CMN001-00
   576001-CMN002-00
   576001-CMN003-00
   576001-CEN001-00
   576001-CEN002-00
   ```
7. Click "Commit Batch" → orders created

**Test PROJECT:**
1. Select Order Type = "PROJECT"
2. Select Project (e.g., TAGO25)
3. Select System Letter
4. Enter quantities
5. Preview and Commit

---

## Summary Checklist

- [ ] Add `btnGetNextBase` button to form
- [ ] Add `cboOrderType_AfterUpdate` event code
- [ ] Add `btnGetNextBase_Click` event code
- [ ] Update `cboProjectSelect_AfterUpdate` if needed
- [ ] Update `Form_Open` to initialize visibility
- [ ] Update `ValidateHeader` for SALES validation
- [ ] Test SALES flow end-to-end
- [ ] Test PROJECT flow still works

---

## Order Number Format Reference

After implementation, both order types will use the unified format:

| Order Type | Format | Example |
|------------|--------|---------|
| SALES | `<BaseToken>-<QualifierCode><SystemLetter><Seq>-<BackorderNo>` | `576001-CMN001-00` |
| PROJECT | `<BaseToken>-<QualifierCode><SystemLetter><Seq>-<BackorderNo>` | `TAGO25-CMN001-00` |

**SALES BaseToken Scheme:**
- 2026: 576000 - 576999
- 2027: 577000 - 577999
- Pattern: `576000 + (year - 2026) * 1000`

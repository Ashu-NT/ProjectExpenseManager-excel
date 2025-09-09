VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_SafetyLine 
   Caption         =   "Safety Equipment"
   ClientHeight    =   3024
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10452
   OleObjectBlob   =   "frm_SafetyLine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_SafetyLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mID As Long
Private mIsDB As Boolean
Private mLoaded As Boolean

Private Sub UserForm_Initialize()
    Dim lo As ListObject, r As Range, arr(), i As Long
    Set lo = GetTable("tblLookups")
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        i = 0
        For Each r In lo.DataBodyRange.rows
            If r.Cells(lo.ListColumns("LookupType").Index).value = "SafetyCategory" Then
                ReDim Preserve arr(i)
                arr(i) = r.Cells(lo.ListColumns("Value").Index).value
                i = i + 1
            End If
        Next r
        If i > 0 Then Me.cmbLineCategory.List = arr
    End If

    ' Format numeric textbox (unit cost)
    FormatNumericTextBox Me.txtUnitCost, 2
    FormatNumericTextBox Me.txtQty, 2

    ' Currency symbol label (create lblUnitCostCur on the form)
    On Error Resume Next
    Me.lblUnitCostCur.Caption = GetSetting("CurrencySymbol", "XAF")
    On Error GoTo 0
End Sub

Private Sub txtUnitCost_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    FormatNumericTextBox Me.txtUnitCost, 2
End Sub
Private Sub txtQty_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    FormatNumericTextBox Me.txtQty, 2
End Sub

Public Sub PrepareEdit(ByVal rowID As Long, ByVal isDB As Boolean)
    mID = rowID
    mIsDB = isDB
    mLoaded = False
End Sub

Private Sub UserForm_Activate()
    If Not mLoaded And mID > 0 Then
        If mIsDB Then LoadFromDBBySafetyID mID Else LoadFromStagingTempID mID
        mLoaded = True
    End If
End Sub

Public Sub LoadFromDBBySafetyID(sID As Long)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblSafety")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        If lr.Range(lo.ListColumns("SafetyID").Index).value = sID Then
            Me.txtLineDate.value = lr.Range(lo.ListColumns("Date").Index).value
            Me.cmbLineCategory.value = lr.Range(lo.ListColumns("CategoryID").Index).value
            Me.txtItemDesc.value = lr.Range(lo.ListColumns("ItemDescription").Index).value
            Me.txtQty.value = lr.Range(lo.ListColumns("Quantity").Index).value
            Me.txtUnitCost.value = lr.Range(lo.ListColumns("UnitCost").Index).value
            Me.txtSupplier.value = lr.Range(lo.ListColumns("Supplier").Index).value
            If ColIndex(lo, "Notes") > 0 Then Me.txtNotes.value = lr.Range(lo.ListColumns("Notes").Index).value
            Exit For
        End If
    Next lr
End Sub

Public Sub LoadFromStagingTempID(tempID As Long)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblStgSafety")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        If lr.Range(lo.ListColumns("TempID").Index).value = tempID Then
            Me.txtLineDate.value = lr.Range(lo.ListColumns("Date").Index).value
            Me.cmbLineCategory.value = lr.Range(lo.ListColumns("CategoryID").Index).value
            Me.txtItemDesc.value = lr.Range(lo.ListColumns("ItemDescription").Index).value
            Me.txtQty.value = lr.Range(lo.ListColumns("Quantity").Index).value
            Me.txtUnitCost.value = lr.Range(lo.ListColumns("UnitCost").Index).value
            Me.txtSupplier.value = lr.Range(lo.ListColumns("Supplier").Index).value
            If ColIndex(lo, "Notes") > 0 Then Me.txtNotes.value = lr.Range(lo.ListColumns("Notes").Index).value
            Exit For
        End If
    Next lr
End Sub

Private Sub btnLineOK_Click()

    On Error GoTo ErrHandler
    Dim userName As String: userName = Environ$("USERNAME")
    Dim loS As ListObject, lr As ListRow, r As Range

    ' Validation (adjust names if your controls differ)
    If Trim(Me.txtLineDate.value) = "" Or Not IsDate(Me.txtLineDate.value) Then
        MsgBox "Date required", vbExclamation: Exit Sub
    End If
    If Trim(Me.cmbLineCategory.value) = "" Then
        MsgBox "Category required", vbExclamation: Exit Sub
    End If
    If Trim(Me.txtItemDesc.value) = "" Then
        MsgBox "Description required", vbExclamation: Exit Sub
    End If

    ' ... numeric validation omitted for brevity, keep yours if needed ...

    If mIsDB Then
        ' Write directly to DB (Add or Update)
        If mID = 0 Then
            ' For Safety form: AddSafetyToDB ... (Material form should call AddMaterialToDB)
            AddSafetyToDB CurrentProjectID, CDate(Me.txtLineDate.value), Trim(Me.cmbLineCategory.value), _
                Trim(Me.txtItemDesc.value), CDbl(Me.txtQty.value), CDbl(Me.txtUnitCost.value), Trim(Me.txtSupplier.value), Trim(Me.txtNotes.value), userName
        Else
            UpdateSafety CLng(mID), CDate(Me.txtLineDate.value), Trim(Me.cmbLineCategory.value), _
                Trim(Me.txtItemDesc.value), CDbl(Me.txtQty.value), CDbl(Me.txtUnitCost.value), Trim(Me.txtSupplier.value), Trim(Me.txtNotes.value), userName
        End If
    Else
        ' Staging branch (Safety): use tblStgSafety. In Material form use tblStgMaterials and Add/UpdateMaterial
        Set loS = GetTable("tblStgSafety")
        If loS Is Nothing Then
            MsgBox "Staging table 'tblStgSafety' not found.", vbCritical: Exit Sub
        End If

        If mID = 0 Then
            Set lr = loS.ListRows.Add
            If ColIndex(loS, "TempID") > 0 Then lr.Range.Cells(1, ColIndex(loS, "TempID")).value = NextID(loS.name, "TempID")
            If ColIndex(loS, "Date") > 0 Then lr.Range.Cells(1, ColIndex(loS, "Date")).value = CDate(Me.txtLineDate.value)
            If ColIndex(loS, "CategoryID") > 0 Then lr.Range.Cells(1, ColIndex(loS, "CategoryID")).value = Trim(Me.cmbLineCategory.value)
            If ColIndex(loS, "ItemDescription") > 0 Then lr.Range.Cells(1, ColIndex(loS, "ItemDescription")).value = Trim(Me.txtItemDesc.value)
            If ColIndex(loS, "Quantity") > 0 Then lr.Range.Cells(1, ColIndex(loS, "Quantity")).value = CDbl(Me.txtQty.value)
            If ColIndex(loS, "UnitCost") > 0 Then lr.Range.Cells(1, ColIndex(loS, "UnitCost")).value = CDbl(Me.txtUnitCost.value)
            If ColIndex(loS, "Supplier") > 0 Then lr.Range.Cells(1, ColIndex(loS, "Supplier")).value = Trim(Me.txtSupplier.value)
            If ColIndex(loS, "Notes") > 0 Then lr.Range.Cells(1, ColIndex(loS, "Notes")).value = Trim(Me.txtNotes.value)
            If ColIndex(loS, "ProjectID") > 0 And CurrentProjectID > 0 Then lr.Range.Cells(1, ColIndex(loS, "ProjectID")).value = CurrentProjectID
        Else
            ' update existing staging row
            Dim foundRow As Range: Set foundRow = Nothing
            If ColIndex(loS, "TempID") > 0 Then
                For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                    If CLng(r.value) = CLng(mID) Then
                        Set foundRow = r.EntireRow
                        Exit For
                    End If
                Next r
            End If
            If foundRow Is Nothing Then
                MsgBox "Staging row not found for TempID=" & mID, vbExclamation: Exit Sub
            Else
                If ColIndex(loS, "Date") > 0 Then foundRow.Cells(1, ColIndex(loS, "Date")).value = CDate(Me.txtLineDate.value)
                If ColIndex(loS, "CategoryID") > 0 Then foundRow.Cells(1, ColIndex(loS, "CategoryID")).value = Trim(Me.cmbLineCategory.value)
                If ColIndex(loS, "ItemDescription") > 0 Then foundRow.Cells(1, ColIndex(loS, "ItemDescription")).value = Trim(Me.txtItemDesc.value)
                If ColIndex(loS, "Quantity") > 0 Then foundRow.Cells(1, ColIndex(loS, "Quantity")).value = CDbl(Me.txtQty.value)
                If ColIndex(loS, "UnitCost") > 0 Then foundRow.Cells(1, ColIndex(loS, "UnitCost")).value = CDbl(Me.txtUnitCost.value)
                If ColIndex(loS, "Supplier") > 0 Then foundRow.Cells(1, ColIndex(loS, "Supplier")).value = Trim(Me.txtSupplier.value)
                If ColIndex(loS, "Notes") > 0 Then foundRow.Cells(1, ColIndex(loS, "Notes")).value = Trim(Me.txtNotes.value)
            End If
        End If
    End If

    ' indicate success to caller and hide the child (do NOT Unload Me here)
    Me.tag = "OK"
    If Not frm_UI Is Nothing Then frm_UI.RefreshStagingLists
    Me.Hide
    Exit Sub

ErrHandler:
    MsgBox "Error saving line: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub


Private Sub btnLineCancel_Click()
    Me.tag = ""
    Me.Hide
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Intercept the window [X] click and hide instead of unloading the form.
    ' This prevents cascade unloads and runtime errors. If user confirms, hide and cancel real close.
    If CloseMode = vbFormControlMenu Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Close this dialog WITHOUT saving changes?" & vbCrLf & _
                     "Choose Yes to close (discard changes), No to return to the form.", _
                     vbYesNo + vbQuestion, "Close dialog")
        If ans = vbYes Then
            ' mark as cancelled so caller knows no save occurred
            Me.tag = ""
            ' hide — does NOT destroy the instance (caller is responsible to Unload)
            Me.Hide
            ' cancel the default close (prevents VBA from unloading the form)
            Cancel = True
        Else
            ' user chose No -> cancel the close and stay on form
            Cancel = True
        End If
    End If
End Sub


' ---------------- validation helpers ----------------
Private Sub txtUnitCost_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = AllowNumericKey(KeyAscii)
End Sub

Private Sub txtQty_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = AllowNumericKey(KeyAscii)
End Sub

Private Sub txtLineDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.txtLineDate.value) <> "" Then
        If Not IsDate(Me.txtLineDate.value) Then
            MsgBox "Please enter a valid date.", vbExclamation
            Cancel = True
            Exit Sub
        End If
        If CDate(Me.txtLineDate.value) > Date Then
            MsgBox "Date cannot be in the future.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
End Sub



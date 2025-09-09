VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_MaterialLIne 
   Caption         =   "Materials"
   ClientHeight    =   3048
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11208
   OleObjectBlob   =   "frm_MaterialLIne.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_MaterialLIne"
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
            If r.Cells(lo.ListColumns("LookupType").Index).value = "MaterialCategory" Then
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
        If mIsDB Then LoadFromDBByMaterialID mID Else LoadFromStagingTempID mID
        mLoaded = True
    End If
End Sub

Public Sub LoadFromDBByMaterialID(matID As Long)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblMaterials")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        If lr.Range(lo.ListColumns("MaterialID").Index).value = matID Then
            Me.txtLineDate.value = lr.Range(lo.ListColumns("Date").Index).value
            Me.cmbLineCategory.value = lr.Range(lo.ListColumns("CategoryID").Index).value
            Me.txtItemDesc.value = lr.Range(lo.ListColumns("ItemDescription").Index).value
            Me.txtQty.value = lr.Range(lo.ListColumns("Quantity").Index).value
            If ColIndex(lo, "Unit") > 0 Then Me.txtUnit.value = lr.Range(lo.ListColumns("Unit").Index).value
            Me.txtUnitCost.value = lr.Range(lo.ListColumns("UnitCost").Index).value
            Me.txtSupplier.value = lr.Range(lo.ListColumns("Supplier").Index).value
            If ColIndex(lo, "Notes") > 0 Then Me.txtNotes.value = lr.Range(lo.ListColumns("Notes").Index).value
            Exit For
        End If
    Next lr
End Sub

Public Sub LoadFromStagingTempID(tempID As Long)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblStgMaterials")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        If lr.Range(lo.ListColumns("TempID").Index).value = tempID Then
            Me.txtLineDate.value = lr.Range(lo.ListColumns("Date").Index).value
            Me.cmbLineCategory.value = lr.Range(lo.ListColumns("CategoryID").Index).value
            Me.txtItemDesc.value = lr.Range(lo.ListColumns("ItemDescription").Index).value
            Me.txtQty.value = lr.Range(lo.ListColumns("Quantity").Index).value
            If ColIndex(lo, "Unit") > 0 Then Me.txtUnit.value = lr.Range(lo.ListColumns("Unit").Index).value
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

    If Not IsDate(Me.txtLineDate.value) Then MsgBox "Date required", vbExclamation: Exit Sub
    If Trim(Me.cmbLineCategory.value) = "" Then MsgBox "Category required", vbExclamation: Exit Sub
    If Trim(Me.txtItemDesc.value) = "" Then MsgBox "Description required", vbExclamation: Exit Sub
    If Not IsNumeric(Me.txtQty.value) Then MsgBox "Quantity must be numeric", vbExclamation: Exit Sub
    If Not IsNumeric(Me.txtUnitCost.value) Then MsgBox "Unit cost must be numeric", vbExclamation: Exit Sub

    If mIsDB Then
        If mID = 0 Then
            AddMaterialToDB CurrentProjectID, CDate(Me.txtLineDate.value), Trim(Me.cmbLineCategory.value), Trim(Me.txtItemDesc.value), _
                CDbl(Me.txtQty.value), Trim(Me.txtUnit.value), CDbl(Me.txtUnitCost.value), Trim(Me.txtSupplier.value), Trim(Me.txtNotes.value), userName
        Else
            UpdateMaterial CLng(mID), CDate(Me.txtLineDate.value), Trim(Me.cmbLineCategory.value), Trim(Me.txtItemDesc.value), _
                CDbl(Me.txtQty.value), Trim(Me.txtUnit.value), CDbl(Me.txtUnitCost.value), Trim(Me.txtSupplier.value), Trim(Me.txtNotes.value), userName
        End If
    Else
        Set loS = GetTable("tblStgMaterials")
        If loS Is Nothing Then MsgBox "tblStgMaterials missing": Exit Sub
        If mID = 0 Then
            Set lr = loS.ListRows.Add
            lr.Range(loS.ListColumns("TempID").Index).value = NextID("tblStgMaterials", "TempID")
            lr.Range(loS.ListColumns("Date").Index).value = CDate(Me.txtLineDate.value)
            lr.Range(loS.ListColumns("CategoryID").Index).value = Trim(Me.cmbLineCategory.value)
            lr.Range(loS.ListColumns("ItemDescription").Index).value = Trim(Me.txtItemDesc.value)
            lr.Range(loS.ListColumns("Quantity").Index).value = CDbl(Me.txtQty.value)
            If ColIndex(loS, "Unit") > 0 Then lr.Range(loS.ListColumns("Unit").Index).value = Trim(Me.txtUnit.value)
            lr.Range(loS.ListColumns("UnitCost").Index).value = CDbl(Me.txtUnitCost.value)
            lr.Range(loS.ListColumns("Supplier").Index).value = Trim(Me.txtSupplier.value)
            If ColIndex(loS, "Notes") > 0 Then lr.Range(loS.ListColumns("Notes").Index).value = Trim(Me.txtNotes.value)
        Else
            For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                If r.value = CLng(mID) Then
                    r.EntireRow.Cells(1, loS.ListColumns("Date").Index).value = CDate(Me.txtLineDate.value)
                    r.EntireRow.Cells(1, loS.ListColumns("CategoryID").Index).value = Trim(Me.cmbLineCategory.value)
                    r.EntireRow.Cells(1, loS.ListColumns("ItemDescription").Index).value = Trim(Me.txtItemDesc.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Quantity").Index).value = CDbl(Me.txtQty.value)
                    If ColIndex(loS, "Unit") > 0 Then r.EntireRow.Cells(1, loS.ListColumns("Unit").Index).value = Trim(Me.txtUnit.value)
                    r.EntireRow.Cells(1, loS.ListColumns("UnitCost").Index).value = CDbl(Me.txtUnitCost.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Supplier").Index).value = Trim(Me.txtSupplier.value)
                    If ColIndex(loS, "Notes") > 0 Then r.EntireRow.Cells(1, loS.ListColumns("Notes").Index).value = Trim(Me.txtNotes.value)
                    Exit For
                End If
            Next r
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



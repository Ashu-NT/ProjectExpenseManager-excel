VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_LogisticsLine 
   Caption         =   "Logistics Line"
   ClientHeight    =   3960
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9516
   OleObjectBlob   =   "frm_LogisticsLine.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_LogisticsLine"
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
            If r.Cells(lo.ListColumns("LookupType").Index).value = "LogisticsCategory" Then
                ReDim Preserve arr(i)
                arr(i) = r.Cells(lo.ListColumns("Value").Index).value
                i = i + 1
            End If
        Next r
        If i > 0 Then Me.cmbLogCategory.List = arr
    End If
    ' Format numerics
    FormatNumericTextBox Me.txtLogAmount, 2
    
    'Currency symbol
    Me.lblLogAmountCur.Caption = GetSetting("CurrencySymbol", "XAF")
End Sub

Private Sub txtLogAmount_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    FormatNumericTextBox Me.txtLogAmount, 2
End Sub


Public Sub PrepareEdit(ByVal rowID As Long, ByVal isDB As Boolean)
    mID = rowID
    mIsDB = isDB
    mLoaded = False
End Sub

Private Sub UserForm_Activate()
    If Not mLoaded And mID > 0 Then
        If mIsDB Then LoadFromDBByLogisticID mID Else LoadFromStagingTempID mID
        mLoaded = True
    End If
End Sub

Public Sub LoadFromDBByLogisticID(logID As Long)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblLogistics")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        If lr.Range(lo.ListColumns("LogisticID").Index).value = logID Then
            Me.txtLogDate.value = lr.Range(lo.ListColumns("Date").Index).value
            Me.cmbLogCategory.value = lr.Range(lo.ListColumns("CategoryID").Index).value
            Me.txtLogDesc.value = lr.Range(lo.ListColumns("Description").Index).value
            Me.txtLogAmount.value = lr.Range(lo.ListColumns("Amount").Index).value
            Me.txtLogVendor.value = lr.Range(lo.ListColumns("Vendor").Index).value
            Exit For
        End If
    Next lr
End Sub

Public Sub LoadFromStagingTempID(tempID As Long)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblStgLogistics")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        If lr.Range(lo.ListColumns("TempID").Index).value = tempID Then
            Me.txtLogDate.value = lr.Range(lo.ListColumns("Date").Index).value
            Me.cmbLogCategory.value = lr.Range(lo.ListColumns("CategoryID").Index).value
            Me.txtLogDesc.value = lr.Range(lo.ListColumns("Description").Index).value
            Me.txtLogAmount.value = lr.Range(lo.ListColumns("Amount").Index).value
            Me.txtLogVendor.value = lr.Range(lo.ListColumns("Vendor").Index).value
            Exit For
        End If
    Next lr
End Sub

Private Sub btnLineOK_Click()
    Dim userName As String: userName = Environ("USERNAME")
    Dim loS As ListObject, lr As ListRow, r As Range

    If Not IsDate(Me.txtLogDate.value) Then MsgBox "Date required", vbExclamation: Exit Sub
    If Trim(Me.cmbLogCategory.value) = "" Then MsgBox "Category required", vbExclamation: Exit Sub
    If Trim(Me.txtLogDesc.value) = "" Then MsgBox "Description required", vbExclamation: Exit Sub
    If Not IsNumeric(Me.txtLogAmount.value) Then MsgBox "Amount must be numeric", vbExclamation: Exit Sub

    If mIsDB Then
        If mID = 0 Then
            AddLogisticToDB CurrentProjectID, CDate(Me.txtLogDate.value), Trim(Me.cmbLogCategory.value), Trim(Me.txtLogDesc.value), CDbl(Me.txtLogAmount.value), Trim(Me.txtLogVendor.value), userName
        Else
            UpdateLogistic CLng(mID), CDate(Me.txtLogDate.value), Trim(Me.cmbLogCategory.value), Trim(Me.txtLogDesc.value), CDbl(Me.txtLogAmount.value), Trim(Me.txtLogVendor.value), userName
        End If
    Else
        Set loS = GetTable("tblStgLogistics")
        If loS Is Nothing Then MsgBox "tblStgLogistics missing": Exit Sub
        If mID = 0 Then
            Set lr = loS.ListRows.Add
            lr.Range(loS.ListColumns("TempID").Index).value = NextID("tblStgLogistics", "TempID")
            lr.Range(loS.ListColumns("Date").Index).value = CDate(Me.txtLogDate.value)
            lr.Range(loS.ListColumns("CategoryID").Index).value = Trim(Me.cmbLogCategory.value)
            lr.Range(loS.ListColumns("Description").Index).value = Trim(Me.txtLogDesc.value)
            lr.Range(loS.ListColumns("Amount").Index).value = CDbl(Me.txtLogAmount.value)
            lr.Range(loS.ListColumns("Vendor").Index).value = Trim(Me.txtLogVendor.value)
        Else
            For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                If r.value = CLng(mID) Then
                    r.EntireRow.Cells(1, loS.ListColumns("Date").Index).value = CDate(Me.txtLogDate.value)
                    r.EntireRow.Cells(1, loS.ListColumns("CategoryID").Index).value = Trim(Me.cmbLogCategory.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Description").Index).value = Trim(Me.txtLogDesc.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Amount").Index).value = CDbl(Me.txtLogAmount.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Vendor").Index).value = Trim(Me.txtLogVendor.value)
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

'----------------------------------Validate date and numerical values----------------------------

Private Sub txtLogAmount_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = AllowNumericKey(KeyAscii)
End Sub

Private Sub txtLogDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.txtLogDate.value) <> "" Then
        If Not IsDate(Me.txtLogDate.value) Then
            MsgBox "Please enter a valid date.", vbExclamation
            Cancel = True
            Exit Sub
        End If
        If CDate(Me.txtLogDate.value) > Date Then
            MsgBox "Date cannot be in the future.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_PaymentLine 
   Caption         =   "Payment Line"
   ClientHeight    =   3216
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11124
   OleObjectBlob   =   "frm_PaymentLine.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_PaymentLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mID As Long
Private mIsDB As Boolean
Private mLoaded As Boolean
Private mRecalcLock As Boolean

Private Sub UserForm_Initialize()
    ' populate workers
    Dim lo As ListObject, r As Range, arr(), i As Long
    Set lo = GetTable("tblWorkers")
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        i = 0
        For Each r In lo.DataBodyRange.rows
            ReDim Preserve arr(i)
            arr(i) = r.Cells(lo.ListColumns("WorkerName").Index).value
            i = i + 1
        Next r
        If i > 0 Then Me.cmbPayWorker.List = arr
    End If
    ' populate payment methods
    Set lo = GetTable("tblLookups")
    If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
        i = 0
        For Each r In lo.DataBodyRange.rows
            If r.Cells(lo.ListColumns("LookupType").Index).value = "PaymentMethod" Then
                ReDim Preserve arr(i)
                arr(i) = r.Cells(lo.ListColumns("Value").Index).value
                i = i + 1
            End If
        Next r
        If i > 0 Then Me.cmbPayMethod.List = arr
    End If
    
    'Format hour
    FormatNumericTextBox Me.txtPayHours, 2
    
    'Currency symbol
    Me.lblPayRateCur.Caption = GetSetting("CurrencySymbol", "XAF")
    Me.lblPayAmountCur.Caption = GetSetting("CurrencySymbol", "XAF")
End Sub

Private Sub txtPayHours_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    FormatNumericTextBox Me.txtPayHours, 2
End Sub


Public Sub PrepareEdit(ByVal rowID As Long, ByVal isDB As Boolean)
    mID = rowID
    mIsDB = isDB
    mLoaded = False
End Sub

Private Sub UserForm_Activate()
    If Not mLoaded And mID > 0 Then
        If mIsDB Then
            LoadFromDBByPaymentID mID
        Else
            LoadFromStagingTempID mID
        End If
        mLoaded = True
    End If
End Sub

Public Sub LoadFromDBByPaymentID(payID As Long)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblPayments")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        If lr.Range(lo.ListColumns("PaymentID").Index).value = payID Then
            Me.txtPayDate.value = lr.Range(lo.ListColumns("DatePaid").Index).value
            Me.cmbPayWorker.value = GetWorkerNameByID(lr.Range(lo.ListColumns("WorkerID").Index).value)
            Me.txtPayHours.value = lr.Range(lo.ListColumns("Hours").Index).value
            Me.txtPayRate.value = lr.Range(lo.ListColumns("Rate").Index).value
            Me.txtPayAmount.value = lr.Range(lo.ListColumns("Amount").Index).value
            Me.cmbPayMethod.value = lr.Range(lo.ListColumns("PaymentMethodID").Index).value
            Me.txtPayNotes.value = lr.Range(lo.ListColumns("Notes").Index).value
            Exit For
        End If
    Next lr
End Sub

Public Sub LoadFromStagingTempID(tempID As Long)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblStgPayments")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For Each lr In lo.ListRows
        If lr.Range(lo.ListColumns("TempID").Index).value = tempID Then
            Me.txtPayDate.value = lr.Range(lo.ListColumns("DatePaid").Index).value
            Me.cmbPayWorker.value = GetWorkerNameByID(lr.Range(lo.ListColumns("WorkerID").Index).value)
            Me.txtPayHours.value = lr.Range(lo.ListColumns("Hours").Index).value
            Me.txtPayRate.value = lr.Range(lo.ListColumns("Rate").Index).value
            Me.txtPayAmount.value = lr.Range(lo.ListColumns("Amount").Index).value
            Me.cmbPayMethod.value = lr.Range(lo.ListColumns("PaymentMethodID").Index).value
            Me.txtPayNotes.value = lr.Range(lo.ListColumns("Notes").Index).value
            Exit For
        End If
    Next lr
End Sub

Private Sub btnLineOK_Click()
    Dim userName As String: userName = Environ("USERNAME")
    Dim loS As ListObject, lr As ListRow, r As Range
    If Trim(Me.cmbPayWorker.value) = "" Then MsgBox "Worker required", vbExclamation: Exit Sub
    If Not IsDate(Me.txtPayDate.value) Then MsgBox "Date required", vbExclamation: Exit Sub
    If Not IsNumeric(Me.txtPayHours.value) Then MsgBox "Hours numeric", vbExclamation: Exit Sub
    If Not IsNumeric(Me.txtPayRate.value) Then MsgBox "Rate numeric", vbExclamation: Exit Sub
    If val(Me.txtPayAmount.value) = 0 Then Me.txtPayAmount.value = CDbl(Me.txtPayHours.value) * CDbl(Me.txtPayRate.value)

    If mIsDB Then
        If mID = 0 Then
            AddPaymentToDB CurrentProjectID, GetWorkerIDByName(Trim(Me.cmbPayWorker.value)), CDate(Me.txtPayDate.value), CDbl(Me.txtPayHours.value), CDbl(Me.txtPayRate.value), CDbl(Me.txtPayAmount.value), Trim(Me.cmbPayMethod.value), Trim(Me.txtPayNotes.value), userName
        Else
            UpdatePayment mID, GetWorkerIDByName(Trim(Me.cmbPayWorker.value)), CDate(Me.txtPayDate.value), CDbl(Me.txtPayHours.value), CDbl(Me.txtPayRate.value), CDbl(Me.txtPayAmount.value), Trim(Me.cmbPayMethod.value), Trim(Me.txtPayNotes.value), userName
        End If
    Else
        Set loS = GetTable("tblStgPayments")
        If loS Is Nothing Then MsgBox "tblStgPayments missing": Exit Sub
        If mID = 0 Then
            Set lr = loS.ListRows.Add
            lr.Range(loS.ListColumns("TempID").Index).value = NextID("tblStgPayments", "TempID")
            lr.Range(loS.ListColumns("WorkerID").Index).value = GetWorkerIDByName(Trim(Me.cmbPayWorker.value))
            lr.Range(loS.ListColumns("DatePaid").Index).value = CDate(Me.txtPayDate.value)
            lr.Range(loS.ListColumns("Hours").Index).value = CDbl(Me.txtPayHours.value)
            lr.Range(loS.ListColumns("Rate").Index).value = CDbl(Me.txtPayRate.value)
            lr.Range(loS.ListColumns("Amount").Index).value = CDbl(Me.txtPayAmount.value)
            lr.Range(loS.ListColumns("PaymentMethodID").Index).value = Trim(Me.cmbPayMethod.value)
            lr.Range(loS.ListColumns("Notes").Index).value = Trim(Me.txtPayNotes.value)
        Else
            For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                If r.value = CLng(mID) Then
                    r.EntireRow.Cells(1, loS.ListColumns("WorkerID").Index).value = GetWorkerIDByName(Trim(Me.cmbPayWorker.value))
                    r.EntireRow.Cells(1, loS.ListColumns("DatePaid").Index).value = CDate(Me.txtPayDate.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Hours").Index).value = CDbl(Me.txtPayHours.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Rate").Index).value = CDbl(Me.txtPayRate.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Amount").Index).value = CDbl(Me.txtPayAmount.value)
                    r.EntireRow.Cells(1, loS.ListColumns("PaymentMethodID").Index).value = Trim(Me.cmbPayMethod.value)
                    r.EntireRow.Cells(1, loS.ListColumns("Notes").Index).value = Trim(Me.txtPayNotes.value)
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

' ------------------- Auto-rate + auto-amount patch -------------------
' Paste this into the frm_PaymentLine code module (below existing code)

' Recalculate the amount from hours and rate
Private Sub RecalcAmount()
    If mRecalcLock Then Exit Sub
    mRecalcLock = True
    On Error GoTo Clean

    Dim h As Double, r As Double, a As Double
    h = ToDbl(Me.txtPayHours.value)
    r = ToDbl(Me.txtPayRate.value)
    a = h * r

    ' Show nicely, but keep numbers numeric in memory
    SetTextSilently Me.txtPayAmount, Fmt2(a)

Clean:
    mRecalcLock = False
End Sub


' When worker changes: populate rate (if available) and recalc
Private Sub cmbPayWorker_Change()
    ' Load rate from Workers table and display it.
    Dim rate As Variant
    rate = GetWorkerRateByName(Trim$(Me.cmbPayWorker.value)) ' your existing lookup
    If IsNumeric(rate) Then
        SetTextSilently Me.txtPayRate, Fmt2(CDbl(rate))
    End If
    RecalcAmount
End Sub

' When hours change: recalc amount
Private Sub txtPayHours_Change()
    RecalcAmount
End Sub

' When rate change: recalc amount
Private Sub txtPayRate_Change()
    RecalcAmount
End Sub

Private Sub SetTextSilently(tb As MSForms.TextBox, ByVal s As String)
    ' Temporarily detach events if you want; here we rely on mRecalcLock instead
    tb.value = s
End Sub


' Helper: robustly lookup a worker's default/hire rate by name
Private Function GetWorkerRateByName(workerName As String) As Double
    On Error GoTo ErrHandler
    Dim lo As ListObject, r As Range
    Dim colRate As Long, colAlt As Variant
    Dim tryCols As Variant
    tryCols = Array("Rate", "HourlyRate", "DefaultRate", "PayRate", "WorkerRate")

    Set lo = GetTable("tblWorkers")
    If lo Is Nothing Then Exit Function

    ' Find which of the candidate columns exists
    colRate = 0
    For Each colAlt In tryCols
        On Error Resume Next
        colRate = lo.ListColumns(CStr(colAlt)).Index
        If Err.Number = 0 Then
            Err.Clear
            Exit For
        Else
            colRate = 0
            Err.Clear
        End If
    Next colAlt

    ' If no rate column found, return 0
    If colRate = 0 Then Exit Function

    ' Search worker name column for a match
    Dim colName As Long
    On Error Resume Next
    colName = lo.ListColumns("WorkerName").Index
    If colName = 0 Then
        ' fallback: try second column if column name differs
        colName = 1
    End If
    On Error GoTo ErrHandler

    If Not lo.DataBodyRange Is Nothing Then
        For Each r In lo.ListColumns(colName).DataBodyRange.rows
            If LCase(Trim(r.value)) = LCase(Trim(workerName)) Then
                GetWorkerRateByName = val(r.EntireRow.Cells(1, colRate).value)
                Exit Function
            End If
        Next r
    End If

    ' not found -> return 0
    Exit Function
ErrHandler:
    ' on error return zero; don't break the form
    GetWorkerRateByName = 0
    Resume Next
End Function
' ------------------- end patch -------------------



' ==========================VALIDATE DATE AND NUMERICAL ENTRIES==================================

Private Sub txtPayHours_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = AllowNumericKey(KeyAscii)
End Sub

Private Sub txtPayRate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = AllowNumericKey(KeyAscii)
End Sub

Private Sub txtPayAmount_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = AllowNumericKey(KeyAscii)
End Sub

Private Sub txtPayDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.txtPayDate.value) <> "" Then
        If Not IsDate(Me.txtPayDate.value) Then
            MsgBox "Please enter a valid date.", vbExclamation
            Cancel = True
            Exit Sub
        End If
        If CDate(Me.txtPayDate.value) > Date Then
            MsgBox "Date cannot be in the future.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

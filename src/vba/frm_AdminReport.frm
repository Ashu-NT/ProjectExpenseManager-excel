VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_AdminReport 
   Caption         =   "Project Reports (Admin)"
   ClientHeight    =   8700
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10008
   OleObjectBlob   =   "frm_AdminReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_AdminReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Initialize()
    ' populate project combo from tblProjects
    Dim lo As ListObject, r As Range, arr(), i As Long
    Set lo = GetTable("tblProjects")
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            ReDim arr(0 To lo.ListRows.Count - 1)
            i = 0
            For Each r In lo.DataBodyRange.rows
                arr(i) = r.Cells(ColIndex(lo, "ProjectName")).value & " [" & r.Cells(ColIndex(lo, "ProjectID")).value & "]"
                i = i + 1
            Next r
            Me.cmbReportProject.List = arr
        End If
    End If

    ' defaults
    Me.chkIncludeConsumables.value = True
    Me.chkIncludePayments.value = True
    Me.chkIncludeLogistics.value = True
    Me.txtFromDate.value = ""
    Me.txtToDate.value = ""
    Me.lblStatus.Caption = ""
End Sub

Private Sub btnLoadProject_Click()
    ' allow typing project ID in square brackets or selecting name
    Dim sel As String, pID As Long
    sel = Trim(Me.cmbReportProject.value)
    pID = ExtractProjectIDFromCombo(sel)
    If pID = 0 Then
        MsgBox "Please select a valid project from the list.", vbExclamation
        Exit Sub
    End If
    Me.lblStatus.Caption = "Loaded ProjectID: " & pID
End Sub

Private Sub btnGenerate_Click()
    Dim sel As String, pID As Long
    sel = Trim(Me.cmbReportProject.value)
    pID = ExtractProjectIDFromCombo(sel)
    If pID = 0 Then
        MsgBox "Please select a project to generate report.", vbExclamation
        Exit Sub
    End If

    Dim dtFrom As Variant, dtTo As Variant
    If Trim(Me.txtFromDate.value) <> "" Then If IsDate(Me.txtFromDate.value) Then dtFrom = CDate(Me.txtFromDate.value)
    If Trim(Me.txtToDate.value) <> "" Then If IsDate(Me.txtToDate.value) Then dtTo = CDate(Me.txtToDate.value)

    Me.lblStatus.Caption = "Generating..."
    Me.Repaint
    PopulatePreview pID, dtFrom, dtTo, Trim(Me.cmbCategoryFilter.value)
    GenerateProjectReport pID, Me.chkIncludeConsumables.value, Me.chkIncludePayments.value, Me.chkIncludeLogistics.value, Me.chkIncludeSafety.value, Me.chkIncludeMaterials.value, dtFrom, dtTo, Trim(Me.cmbCategoryFilter.value)

    Me.lblStatus.Caption = "Generated at " & Format(Now, "yyyy-mm-dd HH:NN")
    AuditWrite "Report", "ProjectReport", pID, Environ$("USERNAME"), "Generated report includeCons=" & Me.chkIncludeConsumables.value & " includePays=" & Me.chkIncludePayments.value & " includeLogs=" & Me.chkIncludeLogistics.value & " includeLogs=" & Me.chkIncludeSafety.value & " includeLogs=" & Me.chkIncludeMaterials.value
End Sub

Private Sub btnExportPDF_Click()
    Dim defaultName As String
    defaultName = "ProjectReport_" & Format(Now, "yyyy-mm-dd_HHNN")
    ExportProjectReportToPDF defaultName
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' --- small helpers used inside form ---
Private Function ExtractProjectIDFromCombo(sel As String) As Long
    On Error Resume Next
    Dim idStart As Long, idEnd As Long, sID As String
    idStart = InStrRev(sel, "[")
    idEnd = InStrRev(sel, "]")
    If idStart > 0 And idEnd > idStart Then
        sID = mID$(sel, idStart + 1, idEnd - idStart - 1)
        ExtractProjectIDFromCombo = CLng(val(sID))
    Else
        ' maybe the user typed only the ID
        If IsNumeric(sel) Then ExtractProjectIDFromCombo = CLng(sel)
    End If
End Function

Private Sub PopulatePreview(pID As Long, dtFrom As Variant, dtTo As Variant, catFilter As String)
    Dim lo As ListObject, r As Range
    Dim dVal As Variant, desc As String, cat As String, amt As Double, wrkID As String
    
    ' Clear existing items
    Me.lstPreviewCons.Clear
    Me.lstPreviewPays.Clear
    Me.lstPreviewLogs.Clear
    Me.lstPreviewSafety.Clear
    Me.lstPreviewMaterials.Clear
    
    ' ==============================
    ' Consumables
    ' ==============================
    If Me.chkIncludeConsumables.value Then
        Set lo = GetTable("tblConsumables")
        If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
            For Each r In lo.DataBodyRange.rows
                If r.Cells(ColIndex(lo, "ProjectID")).value = pID Then
                    dVal = r.Cells(ColIndex(lo, "Date")).value
                    desc = r.Cells(ColIndex(lo, "ItemDescription")).value
                    cat = r.Cells(ColIndex(lo, "CategoryID")).value
                    amt = r.Cells(ColIndex(lo, "TotalCost")).value
                    
                    If PassesDateFilter(dVal, dtFrom, dtTo) Then
                        If (catFilter = "" Or cat = catFilter) Then
                            Me.lstPreviewCons.AddItem
                            With Me.lstPreviewCons
                                .List(.ListCount - 1, 0) = Format(dVal, "yyyy-mm-dd")
                                .List(.ListCount - 1, 1) = cat
                                .List(.ListCount - 1, 2) = desc
                                .List(.ListCount - 1, 3) = Format(amt, "#,##0.00")
                            End With
                        End If
                    End If
                End If
            Next r
        End If
    End If
    
    ' ==============================
    ' Payments
    ' ==============================
    If Me.chkIncludePayments.value Then
        Set lo = GetTable("tblPayments")
        If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
            For Each r In lo.DataBodyRange.rows
                If r.Cells(ColIndex(lo, "ProjectID")).value = pID Then
                    dVal = r.Cells(ColIndex(lo, "DatePaid")).value
                    wrkID = r.Cells(ColIndex(lo, "WorkerID")).value
                    desc = r.Cells(ColIndex(lo, "PaymentMethodID")).value
                    amt = r.Cells(ColIndex(lo, "Amount")).value
                    
                    If PassesDateFilter(dVal, dtFrom, dtTo) Then
                        Me.lstPreviewPays.AddItem
                        With Me.lstPreviewPays
                            .List(.ListCount - 1, 0) = Format(dVal, "yyyy-mm-dd")
                            .List(.ListCount - 1, 1) = wrkID
                            .List(.ListCount - 1, 2) = desc
                            .List(.ListCount - 1, 3) = Format(amt, "#,##0.00")
                        End With
                    End If
                End If
            Next r
        End If
    End If
    
    ' ==============================
    ' Logistics
    ' ==============================
    If Me.chkIncludeLogistics.value Then
        Set lo = GetTable("tblLogistics")
        If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
            For Each r In lo.DataBodyRange.rows
                If r.Cells(ColIndex(lo, "ProjectID")).value = pID Then
                    dVal = r.Cells(ColIndex(lo, "Date")).value
                    desc = r.Cells(ColIndex(lo, "Description")).value
                    cat = r.Cells(ColIndex(lo, "CategoryID")).value
                    amt = r.Cells(ColIndex(lo, "Amount")).value
                    
                    If PassesDateFilter(dVal, dtFrom, dtTo) Then
                        Me.lstPreviewLogs.AddItem
                        With Me.lstPreviewLogs
                            .List(.ListCount - 1, 0) = Format(dVal, "yyyy-mm-dd")
                            .List(.ListCount - 1, 1) = desc
                            .List(.ListCount - 1, 2) = cat
                            .List(.ListCount - 1, 3) = Format(amt, "#,##0.00")
                        End With
                    End If
                End If
            Next r
        End If
    End If
    
    ' ==============================
    ' Safety
    ' ==============================
    If Me.chkIncludeSafety.value Then
        Set lo = GetTable("tblSafety")
        If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
            For Each r In lo.DataBodyRange.rows
                If r.Cells(ColIndex(lo, "ProjectID")).value = pID Then
                    dVal = r.Cells(ColIndex(lo, "Date")).value
                    desc = r.Cells(ColIndex(lo, "ItemDescription")).value
                    cat = r.Cells(ColIndex(lo, "CategoryID")).value
                    amt = r.Cells(ColIndex(lo, "TotalCost")).value
                    
                    If PassesDateFilter(dVal, dtFrom, dtTo) Then
                        Me.lstPreviewSafety.AddItem
                        With Me.lstPreviewSafety
                            .List(.ListCount - 1, 0) = Format(dVal, "yyyy-mm-dd")
                            .List(.ListCount - 1, 1) = desc
                            .List(.ListCount - 1, 2) = cat
                            .List(.ListCount - 1, 3) = Format(amt, "#,##0.00")
                        End With
                    End If
                End If
            Next r
        End If
    End If
    
    ' ==============================
    ' Material
    ' ==============================
    If Me.chkIncludeMaterials.value Then
        Set lo = GetTable("tblMaterials")
        If Not lo Is Nothing And Not lo.DataBodyRange Is Nothing Then
            For Each r In lo.DataBodyRange.rows
                If r.Cells(ColIndex(lo, "ProjectID")).value = pID Then
                    dVal = r.Cells(ColIndex(lo, "Date")).value
                    desc = r.Cells(ColIndex(lo, "ItemDescription")).value
                    cat = r.Cells(ColIndex(lo, "CategoryID")).value
                    amt = r.Cells(ColIndex(lo, "TotalCost")).value
                    
                    If PassesDateFilter(dVal, dtFrom, dtTo) Then
                        Me.lstPreviewMaterials.AddItem
                        With Me.lstPreviewMaterials
                            .List(.ListCount - 1, 0) = Format(dVal, "yyyy-mm-dd")
                            .List(.ListCount - 1, 1) = desc
                            .List(.ListCount - 1, 2) = cat
                            .List(.ListCount - 1, 3) = Format(amt, "#,##0.00")
                        End With
                    End If
                End If
            Next r
        End If
    End If
    
    
End Sub



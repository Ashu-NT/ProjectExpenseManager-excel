VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_UI 
   Caption         =   "Project Entry"
   ClientHeight    =   10452
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12024
   OleObjectBlob   =   "frm_UI.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim loC As ListObject, arr, r As Range, idx As Long

    ' populate companies
    Set loC = GetTable("tblCompanies")
    If Not loC Is Nothing Then
        If Not loC.DataBodyRange Is Nothing Then
            arr = Application.Transpose(loC.ListColumns("CompanyName").DataBodyRange)
            Me.cmbCompany.List = arr
        End If
    End If

    ' populate status from lookups
    Dim loL As ListObject
    Set loL = GetTable("tblLookups")
    If Not loL Is Nothing And Not loL.DataBodyRange Is Nothing Then
        idx = 0
        ReDim arr(0)
        For Each r In loL.DataBodyRange.rows
            If r.Cells(loL.ListColumns("LookupType").Index).value = "ProjectStatus" Then
                ReDim Preserve arr(idx)
                arr(idx) = r.Cells(loL.ListColumns("Value").Index).value
                idx = idx + 1
            End If
        Next r
        If idx > 0 Then Me.cmbStatus.List = arr
    End If

' Configure the results panel (hidden by default)
    With Me.lstProjectResults
        .ColumnCount = 5
        .ColumnWidths = "0;100;200;150;90"
        .ListStyle = fmListStyleOption
        .MultiSelect = fmMultiSelectSingle
    End With
    Me.fraProjectSearch.Visible = False
    
    ' Setup listboxes: include first hidden column for ID
    Me.lstConsumables.ColumnCount = 6
    Me.lstConsumables.ColumnWidths = "0;80;200;60;80;120" ' 0 hides ConsumableID
    Me.lstPayments.ColumnCount = 6
    Me.lstPayments.ColumnWidths = "0;80;200;60;80;120" ' 0 hides PaymentID
    Me.lstLogistics.ColumnCount = 5
    Me.lstLogistics.ColumnWidths = "0;80;220;120;120" ' 0 hides LogisticID
    Me.lstSafety.ColumnCount = 6
    Me.lstSafety.ColumnWidths = "0;80;200;60;80;120"  ' 0 hides SafetyID
    Me.lstMaterials.ColumnCount = 6
    Me.lstMaterials.ColumnWidths = "0;80;200;60;80;120" ' 0 hides MaterialID
    

    Me.lblStatus.Caption = ""
    ' ensure global state
    CurrentProjectID = 0
    Call RefreshStagingLists
    
    ' hide or show settings button on the form based on permissions (if you have btnSettings)
    On Error Resume Next
    Dim btnSettings As MSForms.CommandButton
    Dim btnGenerateReport As MSForms.CommandButton
    
    Set btnSettings = Me.Controls("btnSettings")
    Set btnGenerateReport = Me.Controls("btnGenerateReport")
    If Not btnSettings Is Nothing Then
        If gIsAdminVerified Then
            btnSettings.Visible = True
            btnGenerateReport.Visible = True
        Else
            btnSettings.Visible = False
            btnGenerateReport.Visible = False
        End If
    End If
    On Error GoTo 0
    'Format number
    FormatNumericTextBox Me.txtBudget, 2
    
    'Currency symbol
    Me.lblBudgetCur.Caption = GetSetting("CurrencySymbol", "XAF")
End Sub

Private Sub txtBudget_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    FormatNumericTextBox Me.txtBudget, 2
End Sub

' ----------------- Refresh staging lists (existing behavior) -----------------
Public Sub RefreshStagingLists()
    Dim lo As ListObject, i As Long
    ' Consumables staging
    Set lo = GetTable("tblStgConsumables")
    Me.lstConsumables.Clear
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For i = 1 To lo.ListRows.Count
                Me.lstConsumables.AddItem
                Me.lstConsumables.List(Me.lstConsumables.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("TempID").Index).value
                Me.lstConsumables.List(Me.lstConsumables.ListCount - 1, 1) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Date").Index).value
                Me.lstConsumables.List(Me.lstConsumables.ListCount - 1, 2) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ItemDescription").Index).value
                Me.lstConsumables.List(Me.lstConsumables.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Quantity").Index).value
                Me.lstConsumables.List(Me.lstConsumables.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("UnitCost").Index).value
                Me.lstConsumables.List(Me.lstConsumables.ListCount - 1, 5) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Supplier").Index).value
            Next i
        End If
    End If

    ' Payments staging
    Set lo = GetTable("tblStgPayments")
    Me.lstPayments.Clear
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For i = 1 To lo.ListRows.Count
                Me.lstPayments.AddItem
                Me.lstPayments.List(Me.lstPayments.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("TempID").Index).value
                Me.lstPayments.List(Me.lstPayments.ListCount - 1, 1) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("DatePaid").Index).value
                Me.lstPayments.List(Me.lstPayments.ListCount - 1, 2) = GetWorkerNameByID(lo.DataBodyRange.rows(i).Cells(lo.ListColumns("WorkerID").Index).value)
                Me.lstPayments.List(Me.lstPayments.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Hours").Index).value
                Me.lstPayments.List(Me.lstPayments.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Amount").Index).value
                Me.lstPayments.List(Me.lstPayments.ListCount - 1, 5) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("PaymentMethodID").Index).value
            Next i
        End If
    End If

    ' Logistics staging
    Set lo = GetTable("tblStgLogistics")
    Me.lstLogistics.Clear
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For i = 1 To lo.ListRows.Count
                Me.lstLogistics.AddItem
                Me.lstLogistics.List(Me.lstLogistics.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("TempID").Index).value
                Me.lstLogistics.List(Me.lstLogistics.ListCount - 1, 1) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Date").Index).value
                Me.lstLogistics.List(Me.lstLogistics.ListCount - 1, 2) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Description").Index).value
                Me.lstLogistics.List(Me.lstLogistics.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Amount").Index).value
                Me.lstLogistics.List(Me.lstLogistics.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Vendor").Index).value
            Next i
        End If
    End If
    
    ' Safety staging
    Set lo = GetTable("tblStgSafety")
    Me.lstSafety.Clear
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For i = 1 To lo.ListRows.Count
                Me.lstSafety.AddItem
                Me.lstSafety.List(Me.lstSafety.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("TempID").Index).value
                Me.lstSafety.List(Me.lstSafety.ListCount - 1, 1) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Date").Index).value
                Me.lstSafety.List(Me.lstSafety.ListCount - 1, 2) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ItemDescription").Index).value
                Me.lstSafety.List(Me.lstSafety.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Quantity").Index).value
                Me.lstSafety.List(Me.lstSafety.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("UnitCost").Index).value
                Me.lstSafety.List(Me.lstSafety.ListCount - 1, 5) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Supplier").Index).value
            Next i
        End If
    End If
    
    ' Material staging
    Set lo = GetTable("tblStgMaterials")
    Me.lstMaterials.Clear
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For i = 1 To lo.ListRows.Count
                Me.lstMaterials.AddItem
                Me.lstMaterials.List(Me.lstMaterials.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("TempID").Index).value
                Me.lstMaterials.List(Me.lstMaterials.ListCount - 1, 1) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Date").Index).value
                Me.lstMaterials.List(Me.lstMaterials.ListCount - 1, 2) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ItemDescription").Index).value
                Me.lstMaterials.List(Me.lstMaterials.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Quantity").Index).value
                Me.lstMaterials.List(Me.lstMaterials.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Unit").Index).value
                Me.lstMaterials.List(Me.lstMaterials.ListCount - 1, 5) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("UnitCost").Index).value
            Next i
        End If
    End If
End Sub

' ----------------- Search / Load project -----------------
Private Sub btnSearchProject_Click()
    Dim term As String
    term = Trim(Me.txtSearchProject.text)
    LoadProjectSearchResults term   ' empty term shows all
End Sub

'-----------------Load Selected & Close buttons + double-click to load--------------
Private Sub btnLoadSelected_Click()
    Dim pID As Long
    pID = SelectedProjectID()
    If pID = 0 Then
        MsgBox "Please select a project in the list.", vbExclamation
        Exit Sub
    End If
    LoadProjectByID pID
    Me.fraProjectSearch.Visible = False
End Sub

Private Sub btnHideResults_Click()
    Me.fraProjectSearch.Visible = False
End Sub

Private Sub lstProjectResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If SelectedProjectID() <> 0 Then
        LoadProjectByID SelectedProjectID()
        Me.fraProjectSearch.Visible = False
    End If
End Sub


'------------------- Helper to Load to Result list --------------------

Private Sub LoadProjectSearchResults(ByVal term As String)
    Dim arr As Variant
    Dim rCount As Long
    
    ' Ask modCore for the data array (0-based 2D array or empty array)
    arr = GetProjectsForSearch(term)
    
    ' Check if array empty (two possible empty forms)
    If IsArray(arr) Then
        On Error Resume Next
        rCount = UBound(arr, 1) - LBound(arr, 1) + 1
        If Err.Number <> 0 Then rCount = 0
        On Error GoTo 0
    Else
        rCount = 0
    End If
    
    If rCount = 0 Then
        Me.lstProjectResults.Clear
        Me.fraProjectSearch.Visible = False
        If Trim(term) = "" Then
            MsgBox "No projects found.", vbInformation
        Else
            MsgBox "No matches found for: " & term, vbInformation
        End If
        Exit Sub
    End If
    
    ' Ensure listbox has 5 cols and set suitable widths (first col will be hidden)
    Me.lstProjectResults.ColumnCount = 5
    ' example widths: hide ID, show code/name/start/status
    Me.lstProjectResults.ColumnWidths = "0;100;200;150;90"
    
    ' Assign the array directly to the listbox
    Me.lstProjectResults.Clear
    Me.lstProjectResults.List = arr
    Me.fraProjectSearch.Visible = True
End Sub

'-------------------Load a project row into the Project page-------------------------------

Public Sub LoadProjectByID(ByVal projectID As Long)
    Dim lo As ListObject, r As Range
    Dim loProj As ListObject
    Set loProj = GetTable("tblProjects")
    If loProj Is Nothing Or loProj.DataBodyRange Is Nothing Then
        MsgBox "Projects table not found.", vbCritical: Exit Sub
    End If

    ' Find project row
    Dim found As Boolean: found = False
    For Each r In loProj.DataBodyRange.rows
        If r.Cells(1, loProj.ListColumns("ProjectID").Index).value = projectID Then
            found = True
            Me.txtProjectName.value = r.Cells(1, loProj.ListColumns("ProjectName").Index).value
            Me.txtProjectCode.value = r.Cells(1, loProj.ListColumns("ProjectCode").Index).value
            Me.cmbCompany.value = GetCompanyNameByID(r.Cells(1, loProj.ListColumns("CompanyID").Index).value)
            Me.dtStart.value = r.Cells(1, loProj.ListColumns("StartDate").Index).value
            Me.dtEnd.value = r.Cells(1, loProj.ListColumns("EndDate").Index).value
            Me.txtBudget.value = r.Cells(1, loProj.ListColumns("Budget").Index).value
            Me.txtManager.value = r.Cells(1, loProj.ListColumns("ProjectManager").Index).value
            Me.cmbStatus.value = r.Cells(1, loProj.ListColumns("Status").Index).value
            Me.txtNotes.value = r.Cells(1, loProj.ListColumns("Notes").Index).value
            Exit For
        End If
    Next r

    If Not found Then
        MsgBox "Project not found: ID " & projectID, vbExclamation
        Exit Sub
    End If

    ' Set global current project
    CurrentProjectID = projectID

    ' --- Copy DB child rows into staging tables ---
    ' Clear staging first (so the staging becomes the authoritative edit set)
    ClearTable GetTable("tblStgConsumables")
    ClearTable GetTable("tblStgPayments")
    ClearTable GetTable("tblStgLogistics")
    ClearTable GetTable("tblStgSafety")
    ClearTable GetTable("tblStgMaterials")

    ' Copy each child table rows having ProjectID = projectID
    CopyProjectRowsToStaging "tblConsumables", "tblStgConsumables", projectID, Array("Date", "CategoryID", "ItemDescription", "Quantity", "UnitCost", "Supplier")
    CopyProjectRowsToStaging "tblPayments", "tblStgPayments", projectID, Array("WorkerID", "DatePaid", "Hours", "Rate", "Amount", "PaymentMethodID", "Notes")
    CopyProjectRowsToStaging "tblLogistics", "tblStgLogistics", projectID, Array("Date", "CategoryID", "Description", "Amount", "Vendor")
    CopyProjectRowsToStaging "tblSafety", "tblStgSafety", projectID, Array("Date", "CategoryID", "ItemDescription", "Quantity", "UnitCost", "Supplier", "Notes")
    CopyProjectRowsToStaging "tblMaterials", "tblStgMaterials", projectID, Array("Date", "CategoryID", "ItemDescription", "Quantity", "UnitCost", "Supplier", "Notes", "Unit")

    ' Now refresh listboxes from staging so they reflect the staging content
    RefreshStagingLists

End Sub


Private Function SelectedProjectID() As Long
    If Me.lstProjectResults.ListIndex < 0 Then
        SelectedProjectID = 0
    Else
        SelectedProjectID = CLng(Me.lstProjectResults.List(Me.lstProjectResults.ListIndex, 0))
    End If
End Function

' Load project header + lists
Public Sub LoadProjectIntoForm(projectID As Long)
    Dim loP As ListObject, r As Range, projRow As Range
    Set loP = GetTable("tblProjects")
    If loP Is Nothing Or loP.DataBodyRange Is Nothing Then MsgBox "tblProjects missing": Exit Sub

    For Each r In loP.ListColumns("ProjectID").DataBodyRange.rows
        If r.value = projectID Then
            Set projRow = r.EntireRow
            Exit For
        End If
    Next r
    If projRow Is Nothing Then MsgBox "Project not found: " & projectID: Exit Sub

    With Me
        .txtProjectCode.value = projRow.Cells(1, loP.ListColumns("ProjectCode").Index).value
        .txtProjectName.value = projRow.Cells(1, loP.ListColumns("ProjectName").Index).value
        Dim compID As Variant
        compID = projRow.Cells(1, loP.ListColumns("CompanyID").Index).value
        If Not IsEmpty(compID) Then .cmbCompany.value = GetCompanyNameByID(compID) Else .cmbCompany.value = ""
        .dtStart.value = projRow.Cells(1, loP.ListColumns("StartDate").Index).value
        .dtEnd.value = projRow.Cells(1, loP.ListColumns("EndDate").Index).value
        .txtBudget.value = projRow.Cells(1, loP.ListColumns("Budget").Index).value
        .txtManager.value = projRow.Cells(1, loP.ListColumns("ProjectManager").Index).value
        .cmbStatus.value = projRow.Cells(1, loP.ListColumns("Status").Index).value
        .txtNotes.value = projRow.Cells(1, loP.ListColumns("Notes").Index).value
    End With

    CurrentProjectID = projectID

    ' Populate DB lists
    PopulateConsumablesListbox Me, projectID
    PopulatePaymentsListbox Me, projectID
    PopulateLogisticsListbox Me, projectID
    PopulateSafetyListbox Me, projectID
    PopulateMaterialsListbox Me, projectID

    Me.lblStatus.Caption = "Loaded ProjectID: " & projectID
End Sub

' ----------------- Save new project (keeps staging commit behavior) -----------------
Private Sub btnSaveProject_Click()
    ' If editing an existing loaded project, just update header
    If CurrentProjectID > 0 Then
        Call btnUpdateProject_Click
        Exit Sub
    End If

    ' Validation
    If Trim(Me.txtProjectCode.value) = "" Then MsgBox "Project code required.", vbExclamation: Exit Sub
    If Trim(Me.cmbCompany.value) = "" Then MsgBox "Company required.", vbExclamation: Exit Sub
    If Not IsDate(Me.dtStart.value) Then MsgBox "Start date required.", vbExclamation: Exit Sub

    ' Check for duplicate by ProjectCode
    Dim existingID As Long
    existingID = FindProjectByCode(Trim(Me.txtProjectCode.value))
    If existingID > 0 Then
        If MsgBox("A project with this ProjectCode already exists (ID " & existingID & ")." & vbCrLf & _
                  "Click Yes to load the existing project (no duplicate will be created)." & vbCrLf & _
                  "Click No to cancel save.", vbYesNo + vbQuestion, "Duplicate ProjectCode") = vbYes Then
            LoadProjectIntoForm existingID
            Exit Sub
        Else
            Exit Sub
        End If
    End If

    ' Create new project (this routine should set CurrentProjectID to the new ID)
    Dim newID As Long
    newID = CreateProjectFromForm(Me) ' returns new ProjectID and sets CurrentProjectID
    If newID = 0 Then Exit Sub

    ' Decide commit behavior based on settings
    Dim autoCommit As Boolean
    autoCommit = (UCase(GetSetting("AutoCommitOnSave", "FALSE")) = "TRUE")

    ' If staging has rows, commit or prompt according to settings
    If HasAnyStagingRows() Then
        ' Safety cap
        If Not CheckMaxRowsBeforeCommit() Then
            Me.lblStatus.Caption = "Save cancelled (exceeded MaxRowsPerCommit)"
            Exit Sub
        End If

        If autoCommit Then
            Dim commitSummaryAC As String
            commitSummaryAC = CommitStagingToDB(CurrentProjectID, Environ$("USERNAME"))
            MsgBox "Project created and committed." & vbCrLf & commitSummaryAC, vbInformation
        Else
            If ConfirmStagingCommit(False) Then   ' new project: duplicate warning not necessary
                Dim commitSummary As String
                commitSummary = CommitStagingToDB(CurrentProjectID, Environ$("USERNAME"))
                MsgBox "Project created and committed." & vbCrLf & commitSummary, vbInformation
            Else
                MsgBox "Project created. Staging not committed.", vbInformation
            End If
        End If
    Else
        MsgBox "Project created. No staging rows to commit.", vbInformation
    End If

    ' Refresh DB lists / staging lists
    On Error Resume Next
    RefreshStagingLists
    PopulateConsumablesListbox Me, CurrentProjectID
    PopulatePaymentsListbox Me, CurrentProjectID
    PopulateLogisticsListbox Me, CurrentProjectID
    PopulateSafetyListbox Me, CurrentProjectID
    PopulateMaterialsListbox Me, CurrentProjectID
    On Error GoTo 0

    Me.lblStatus.Caption = "Saved as ProjectID: " & CurrentProjectID
End Sub



' ----------------- Update existing project header -----------------

Private Sub btnUpdateProject_Click()
    Dim loProj As ListObject, rProj As Range
    Dim updatedID As Long
    Dim userName As String: userName = Environ$("USERNAME")
    Dim found As Boolean: found = False

    If CurrentProjectID <= 0 Then
        MsgBox "No project loaded to update.", vbExclamation
        Exit Sub
    End If
    updatedID = CurrentProjectID

    ' Validate header (reuse your ValidateProjectForm if available)
    If Not ValidateProjectForm(Me) Then Exit Sub

    ' Locate Projects table and update the header row
    Set loProj = GetTable("tblProjects")
    If loProj Is Nothing Or loProj.DataBodyRange Is Nothing Then
        MsgBox "Projects table not found.", vbCritical: Exit Sub
    End If

    For Each rProj In loProj.DataBodyRange.rows
        If rProj.Cells(1, loProj.ListColumns("ProjectID").Index).value = updatedID Then
            ' Update header columns (safe writes)
            rProj.Cells(1, loProj.ListColumns("ProjectCode").Index).value = Trim(Me.txtProjectCode.value)
            rProj.Cells(1, loProj.ListColumns("ProjectName").Index).value = Trim(Me.txtProjectName.value)
            rProj.Cells(1, loProj.ListColumns("CompanyID").Index).value = GetCompanyIDByName(Trim(Me.cmbCompany.value))
            rProj.Cells(1, loProj.ListColumns("StartDate").Index).value = CDate(Me.dtStart.value)
            If Trim(Me.dtEnd.value) <> "" Then
                rProj.Cells(1, loProj.ListColumns("EndDate").Index).value = CDate(Me.dtEnd.value)
            Else
                rProj.Cells(1, loProj.ListColumns("EndDate").Index).ClearContents
            End If
            If IsNumeric(Me.txtBudget.value) Then
                rProj.Cells(1, loProj.ListColumns("Budget").Index).value = CDbl(Me.txtBudget.value)
            Else
                MsgBox "Please enter a valid numeric value for Budget.", vbExclamation
                Exit Sub
            End If
            'rProj.Cells(1, loProj.ListColumns("Budget").Index).value = val(Me.txtBudget.value)
            rProj.Cells(1, loProj.ListColumns("ProjectManager").Index).value = Trim(Me.txtManager.value)
            rProj.Cells(1, loProj.ListColumns("Status").Index).value = Trim(Me.cmbStatus.value)
            rProj.Cells(1, loProj.ListColumns("Notes").Index).value = Trim(Me.txtNotes.value)
            found = True
            Exit For
        End If
    Next rProj

    If Not found Then
        MsgBox "Could not locate ProjectID=" & updatedID & " in tblProjects.", vbExclamation
        Exit Sub
    End If

    ' Audit header update
    AuditWrite "Update", "tblProjects", updatedID, userName, "Project header updated via UI"

    ' --- Now commit staging rows if present ---
    If HasAnyStagingRows() Then

        ' ----- NEW: RequireClearBeforeCommit enforcement -----
        ' If the settings require clearing before commit, offer to clear now (abort update)
        If UCase(GetSetting("RequireClearBeforeCommit", "TRUE")) = "TRUE" Then
            Dim ansClear As VbMsgBoxResult
            ansClear = MsgBox( _
               "Staging currently contains items. If these include previously loaded DB rows that were not cleared, committing may create duplicates." & vbCrLf & vbCrLf & _
               "Click Yes to CLEAR staging now (you will need to re-add new items). Click No to continue with commit.", _
               vbYesNo + vbExclamation, "Clear staging before commit?")
            If ansClear = vbYes Then
                ' Clear staging tables and refresh UI listboxes; abort update so user can re-add items
                ClearTable GetTable("tblStgConsumables")
                ClearTable GetTable("tblStgPayments")
                ClearTable GetTable("tblStgLogistics")
                ClearTable GetTable("tblStgSafety")
                ClearTable GetTable("tblStgMaterials")
                RefreshStagingLists
                Me.lblStatus.Caption = "Staging cleared. Re-add new items and Save."
                MsgBox "Staging cleared. Please add new items and then Save/Update again.", vbInformation
                Exit Sub
            End If
            ' If ansClear = vbNo, continue to confirmation/commit path below
        End If
        ' ----- END RequireClearBeforeCommit enforcement -----

        ' Optional: enforce MaxRowsPerCommit safety cap
        If Not CheckMaxRowsBeforeCommit() Then
            ' User chose not to proceed
            Me.lblStatus.Caption = "Commit cancelled (exceeded MaxRowsPerCommit)"
            Exit Sub
        End If

        ' Ask user to confirm commit (existing routine)
        If Not ConfirmStagingCommit(True) Then
            ' user cancelled commit; still show header updated
            Me.lblStatus.Caption = "Project header updated (no commit)"
            MsgBox "Project header updated. Staging commit cancelled by user.", vbInformation
            Exit Sub
        End If

        ' Call commit and show summary
        Dim commitSummary As String
        On Error GoTo CommitErr
        commitSummary = CommitStagingToDB(updatedID, userName)
        On Error GoTo 0

        ' Refresh UI lists after commit
        RefreshStagingLists
        PopulateConsumablesListbox Me, updatedID
        PopulatePaymentsListbox Me, updatedID
        PopulateLogisticsListbox Me, updatedID
        PopulateSafetyListbox Me, updatedID
        PopulateMaterialsListbox Me, updatedID

        Me.lblStatus.Caption = "Project updated (ProjectID: " & updatedID & ")"
        MsgBox "Project updated (ProjectID = " & updatedID & ")." & vbCrLf & commitSummary, vbInformation
    Else
        ' No staging rows: only header updated
        Me.lblStatus.Caption = "Project header updated (no staging to commit)"
        MsgBox "Project header updated (no staging rows found).", vbInformation
    End If

    Exit Sub

CommitErr:
    MsgBox "Error while committing staging rows: " & Err.Description, vbCritical
End Sub



' ----------------- Advance status -----------------
Private Sub btnAdvanceStatus_Click()
    If CurrentProjectID = 0 Then MsgBox "Load a project first.", vbExclamation: Exit Sub
    Dim seq As Variant, i As Long, cur As String, nextStatus As String
    seq = Array("Planned", "Active", "Completed", "On Hold")
    cur = Trim(Me.cmbStatus.value)
    nextStatus = cur
    For i = LBound(seq) To UBound(seq) - 1
        If seq(i) = cur Then
            nextStatus = seq(i + 1)
            Exit For
        End If
    Next i
    If nextStatus = cur Then nextStatus = seq(0)
    Me.cmbStatus.value = nextStatus
    Call btnUpdateProject_Click
    Me.lblStatus.Caption = "Status advanced to: " & nextStatus
End Sub

Private Sub btnSettings_Click()
    ' Use the secure wrapper so permission checks are applied
    On Error Resume Next
    frm_Settings.Show
    On Error GoTo 0
End Sub


Private Sub btnGenerateReport_Click()
    On Error Resume Next
    ShowAdminReportForm
    On Error GoTo 0
End Sub

' ----------------- Add / Edit / Remove handlers (examples for consumables) -----------------
Private Sub btnAddConsumable_Click()
    frm_ConsumableLine.tag = ""  ' add
    frm_ConsumableLine.Show vbModal
    ' small form already writes into tblStgConsumables on OK
    RefreshStagingLists
End Sub


Private Sub btnEditConsumable_Click()
    Dim idx As Long, id As Long
    If Me.lstConsumables.ListCount = 0 Then MsgBox "No consumables", vbExclamation: Exit Sub
    idx = Me.lstConsumables.ListIndex
    If idx < 0 Then MsgBox "Select a consumable", vbExclamation: Exit Sub

    id = GetSelectedIDFromListbox(Me.lstConsumables)
    If id = 0 Then MsgBox "Could not determine selected ID", vbExclamation: Exit Sub

    Dim isDB As Boolean
    ' If it exists in staging, it's staging (even when a project is loaded)
    If IsIDInStaging("tblStgConsumables", id) Then
        isDB = False
    Else
        isDB = True
    End If

    Dim f As New frm_ConsumableLine
    f.PrepareEdit id, isDB
    f.Show vbModal
    If f.tag = "OK" Then
        RefreshStagingLists
        'If CurrentProjectID > 0 Then PopulateConsumablesListbox Me, CurrentProjectID Else RefreshStagingLists
    End If
    Unload f
    Set f = Nothing
End Sub


Private Sub btnRemoveConsumable_Click()
    Dim idx As Long, id As Long
    idx = Me.lstConsumables.ListIndex
    If idx < 0 Then MsgBox "Select a consumable", vbExclamation: Exit Sub
    id = CLng(Me.lstConsumables.List(idx, 0))
    If MsgBox("Delete selected line?", vbYesNo + vbQuestion) = vbYes Then
        If CurrentProjectID > 0 Then
            DeleteConsumable id, Environ$("USERNAME")
            PopulateConsumablesListbox Me, CurrentProjectID
        Else
            ' staging delete
            Dim loS As ListObject, r As Range
            Set loS = GetTable("tblStgConsumables")
            If loS Is Nothing Then Exit Sub
            For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                If r.value = id Then r.EntireRow.Delete: Exit For
            Next r
            RefreshStagingLists
        End If
    End If
End Sub

' ----------------- Payments handlers -----------------
' Payments
Private Sub btnAddPayment_Click()
  
    frm_PaymentLine.tag = ""
    frm_PaymentLine.Show vbModal
    RefreshStagingLists
    
End Sub


Private Sub btnEditPayment_Click()
    Dim idx As Long, id As Long
    If Me.lstPayments.ListCount = 0 Then MsgBox "No payments", vbExclamation: Exit Sub
    idx = Me.lstPayments.ListIndex
    If idx < 0 Then MsgBox "Select a payment", vbExclamation: Exit Sub

    id = GetSelectedIDFromListbox(Me.lstPayments)
    If id = 0 Then MsgBox "Could not determine selected ID", vbExclamation: Exit Sub

    Dim isDB As Boolean
    If IsIDInStaging("tblStgPayments", id) Then
        isDB = False
    Else
        isDB = True
    End If

    Dim f As New frm_PaymentLine
    f.PrepareEdit id, isDB
    f.Show vbModal
    If f.tag = "OK" Then
        RefreshStagingLists
        'If CurrentProjectID > 0 Then PopulatePaymentsListbox Me, CurrentProjectID Else RefreshStagingLists
    End If
    Unload f
    Set f = Nothing
End Sub


Private Sub btnRemovePayment_Click()
    Dim idx As Long, id As Long
    idx = Me.lstPayments.ListIndex
    If idx < 0 Then MsgBox "Select a payment", vbExclamation: Exit Sub
    id = CLng(Me.lstPayments.List(idx, 0))
    If MsgBox("Delete selected payment?", vbYesNo + vbQuestion) = vbYes Then
        If CurrentProjectID > 0 Then
            DeletePayment id, Environ$("USERNAME")
            PopulatePaymentsListbox Me, CurrentProjectID
        Else
            Dim loS As ListObject, r As Range
            Set loS = GetTable("tblStgPayments")
            If loS Is Nothing Then Exit Sub
            For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                If r.value = id Then r.EntireRow.Delete: Exit For
            Next r
            RefreshStagingLists
        End If
    End If
End Sub

' ----------------- Logistics handlers -----------------
' Logistics
Private Sub btnAddLogistic_Click()

    frm_LogisticsLine.tag = ""
    frm_LogisticsLine.Show vbModal

    RefreshStagingLists
End Sub

Private Sub btnEditLogistic_Click()
    Dim idx As Long, id As Long
    If Me.lstLogistics.ListCount = 0 Then MsgBox "No logistics items", vbExclamation: Exit Sub
    idx = Me.lstLogistics.ListIndex
    If idx < 0 Then MsgBox "Select a logistics item", vbExclamation: Exit Sub

    id = GetSelectedIDFromListbox(Me.lstLogistics)
    If id = 0 Then MsgBox "Could not determine selected ID", vbExclamation: Exit Sub

    Dim isDB As Boolean
    If IsIDInStaging("tblStgLogistics", id) Then
        isDB = False
    Else
        isDB = True
    End If

    Dim f As New frm_LogisticsLine
    f.PrepareEdit id, isDB
    f.Show vbModal
    If f.tag = "OK" Then
        RefreshStagingLists
        'If CurrentProjectID > 0 Then PopulateLogisticsListbox Me, CurrentProjectID Else RefreshStagingLists
    End If
    Unload f
    Set f = Nothing
End Sub


Private Sub btnRemoveLogistic_Click()
    Dim idx As Long, id As Long
    idx = Me.lstLogistics.ListIndex
    If idx < 0 Then MsgBox "Select a logistics line", vbExclamation: Exit Sub
    id = CLng(Me.lstLogistics.List(idx, 0))
    If MsgBox("Delete selected line?", vbYesNo + vbQuestion) = vbYes Then
        If CurrentProjectID > 0 Then
            DeleteLogistic id, Environ$("USERNAME")
            PopulateLogisticsListbox Me, CurrentProjectID
        Else
            Dim loS As ListObject, r As Range
            Set loS = GetTable("tblStgLogistics")
            If loS Is Nothing Then Exit Sub
            For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                If r.value = id Then r.EntireRow.Delete: Exit For
            Next r
            RefreshStagingLists
        End If
    End If
End Sub

' ----------------- Safety handlers -----------------
' Add Safety — use a NEW instance and let parent unload it
Private Sub btnAddSafety_Click()
    Dim f As New frm_SafetyLine
    f.PrepareEdit 0, False      ' new staging record
    f.Show vbModal
    If f.tag = "OK" Then RefreshStagingLists
    Unload f
    Set f = Nothing
End Sub

' Edit Safety
Private Sub btnEditSafety_Click()
    Dim idx As Long, id As Long
    If Me.lstSafety.ListCount = 0 Then MsgBox "No safety items", vbExclamation: Exit Sub
    idx = Me.lstSafety.ListIndex
    If idx < 0 Then MsgBox "Select a safety item", vbExclamation: Exit Sub

    id = GetSelectedIDFromListbox(Me.lstSafety)
    If id = 0 Then MsgBox "Could not determine selected ID", vbExclamation: Exit Sub

    Dim isDB As Boolean
    If IsIDInStaging("tblStgSafety", id) Then
        isDB = False
    Else
        isDB = True
    End If

    Dim f As New frm_SafetyLine
    f.PrepareEdit id, isDB
    f.Show vbModal
    If f.tag = "OK" Then
        RefreshStagingLists
        'If CurrentProjectID > 0 Then PopulateSafetyListbox Me, CurrentProjectID Else RefreshStagingLists
    End If
    Unload f
    Set f = Nothing
End Sub



Private Sub btnRemoveSafety_Click()
    Dim idx As Long, id As Long
    idx = Me.lstSafety.ListIndex
    If idx < 0 Then MsgBox "Select a safety item", vbExclamation: Exit Sub
    id = CLng(Me.lstSafety.List(idx, 0))
    If MsgBox("Delete selected line?", vbYesNo + vbQuestion) = vbYes Then
        If CurrentProjectID > 0 Then
            DeleteSafety id, Environ$("USERNAME")
            PopulateSafetyListbox Me, CurrentProjectID
        Else
            ' staging delete
            Dim loS As ListObject, r As Range
            Set loS = GetTable("tblStgSafety")
            If loS Is Nothing Then Exit Sub
            If Not loS.DataBodyRange Is Nothing Then
                For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                    If val(r.value) = id Then
                        r.EntireRow.Delete
                        Exit For
                    End If
                Next r
            End If
            RefreshStagingLists
        End If
    End If
End Sub

Private Sub btnAddMaterial_Click()
    Dim f As New frm_MaterialLIne
    f.PrepareEdit 0, False
    f.Show vbModal
    If f.tag = "OK" Then RefreshStagingLists
    Unload f
    Set f = Nothing
End Sub

Private Sub btnEditMaterial_Click()
    Dim idx As Long, id As Long
    If Me.lstMaterials.ListCount = 0 Then MsgBox "No material items", vbExclamation: Exit Sub
    idx = Me.lstMaterials.ListIndex
    If idx < 0 Then MsgBox "Select a material item", vbExclamation: Exit Sub

    id = GetSelectedIDFromListbox(Me.lstMaterials)
    If id = 0 Then MsgBox "Could not determine selected ID", vbExclamation: Exit Sub

    Dim isDB As Boolean
    If IsIDInStaging("tblStgMaterials", id) Then
        isDB = False
    Else
        isDB = True
    End If

    Dim f As New frm_MaterialLIne

    'MsgBox "DEBUG: parent" & vbCrLf & _
    '   "Selected ID = " & id & vbCrLf & _
    '   "isDB = " & isDB, vbInformation

    f.PrepareEdit id, isDB
    f.Show vbModal
    If f.tag = "OK" Then
        RefreshStagingLists
        'If CurrentProjectID > 0 Then PopulateMaterialsListbox Me, CurrentProjectID Else RefreshStagingLists
    End If
    Unload f
    Set f = Nothing
End Sub



Private Sub btnRemoveMaterial_Click()
    Dim idx As Long, id As Long
    idx = Me.lstMaterials.ListIndex
    If idx < 0 Then MsgBox "Select a material", vbExclamation: Exit Sub
    id = CLng(Me.lstMaterials.List(idx, 0))
    If MsgBox("Delete selected line?", vbYesNo + vbQuestion) = vbYes Then
        If CurrentProjectID > 0 Then
            DeleteMaterial id, Environ$("USERNAME")
            PopulateMaterialsListbox Me, CurrentProjectID
        Else
            ' staging delete
            Dim loS As ListObject, r As Range
            Set loS = GetTable("tblStgMaterials")
            If loS Is Nothing Then Exit Sub
            If Not loS.DataBodyRange Is Nothing Then
                For Each r In loS.ListColumns("TempID").DataBodyRange.rows
                    If val(r.value) = id Then
                        r.EntireRow.Delete
                        Exit For
                    End If
                Next r
            End If
            RefreshStagingLists
        End If
    End If
End Sub


' ----------------- New / Cancel -----------------
Private Sub btnNewProject_Click()
    CurrentProjectID = 0
    Me.txtProjectCode.value = ""
    Me.txtProjectName.value = ""
    Me.cmbCompany.value = ""
    Me.dtStart.value = ""
    Me.dtEnd.value = ""
    Me.txtBudget.value = ""
    Me.txtManager.value = ""
    Me.cmbStatus.value = ""
    Me.txtNotes.value = ""
    Me.txtSearchProject.value = ""
    RefreshStagingLists
    Me.lblStatus.Caption = "New project (unsaved)"
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


' ----------------- Clear staging button -----------------
Private Sub btnClearStaging_Click()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("This will clear all currently loaded items in the staging tables." & vbCrLf & _
                 "Use this when you want to add NEW items only (prevents duplicates)." & vbCrLf & _
                 "Proceed?", vbYesNo + vbQuestion, "Clear staging?")
    If ans <> vbYes Then Exit Sub

    ' Clear staging tables (structure remains)
    ClearTable GetTable("tblStgConsumables")
    ClearTable GetTable("tblStgPayments")
    ClearTable GetTable("tblStgLogistics")
    ClearTable GetTable("tblStgSafety")
    ClearTable GetTable("tblStgMaterials")

    ' Refresh UI listboxes
    On Error Resume Next
    RefreshStagingLists
    On Error GoTo 0

    MsgBox "Staging tables cleared. You may now add new items.", vbInformation
End Sub




' ---------------- btnDeleteProject_Click ----------------
Private Sub btnDeleteProject_Click()
    Dim curID As Long
    curID = CurrentProjectID

    If curID <= 0 Then
        MsgBox "No project is currently loaded. Load a project before attempting deletion.", vbExclamation
        Exit Sub
    End If

    ' Basic info to show in prompts
    Dim pCode As String, pName As String
    pCode = Trim(Me.txtProjectCode.value)
    pName = Trim(Me.txtProjectName.value)

    ' 1st confirmation: explicit Yes/No
    If MsgBox("You are about to PERMANENTLY DELETE the project:" & vbCrLf & _
              "ProjectID: " & curID & vbCrLf & "ProjectCode: " & pCode & vbCrLf & "ProjectName: " & pName & _
              vbCrLf & vbCrLf & "This will delete all linked consumables, payments and logistics." & vbCrLf & _
              "Do you want to continue?", vbYesNo + vbCritical, "Confirm Delete") <> vbYes Then
        Exit Sub
    End If

    ' Permission check: allowed usernames or admin password
    If Not CanUserDeleteProject() Then
        MsgBox "You are not authorized to delete projects.", vbExclamation
        Exit Sub
    End If

    ' 2nd confirmation: require typing the ProjectCode exactly
    Dim typed As String
    typed = InputBox("Type the Project Code to CONFIRM deletion:" & vbCrLf & "(case-sensitive)", "Confirm deletion by typing ProjectCode")
    If StrPtr(typed) = 0 Then Exit Sub ' cancelled
    If Trim(typed) <> pCode Then
        MsgBox "ProjectCode does not match. Deletion cancelled.", vbExclamation
        Exit Sub
    End If

    ' Optional backup before delete (respect tblSettings BackupBeforeCommit)
    If UCase(GetSetting("BackupBeforeCommit", "TRUE")) = "TRUE" Then
        On Error Resume Next
        CreateBackupCopy "delete_project_" & curID
        On Error GoTo 0
    End If

    ' Perform deletion
    Dim ok As Boolean
    ok = DeleteProjectByID(curID)

    If ok Then
        MsgBox "Project and related records deleted successfully (ProjectID = " & curID & ").", vbInformation
        ' Clear UI & reset state
        CurrentProjectID = 0
        ClearProjectForm Me
        RefreshStagingLists
        ' Refresh any DB lists you display
        PopulateConsumablesListbox Me, 0
        PopulatePaymentsListbox Me, 0
        PopulateLogisticsListbox Me, 0
        PopulateSafetyListbox Me, 0
        PopulateMaterialsListbox Me, 0
        
        Me.lblStatus.Caption = "Deleted ProjectID: " & curID
    Else
        MsgBox "Deletion failed — see audit or VBA Immediate for details.", vbCritical
    End If
End Sub




'-------------------------------VALIDATE DATE AND NUMERICAL VALUES=========================

Private Sub txtBudget_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = AllowNumericKey(KeyAscii)
End Sub

Private Sub dtEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.dtEnd.value) <> "" Then
        If Not IsDate(Me.dtEnd.value) Then
            MsgBox "Please enter a valid date.", vbExclamation
            Cancel = True
            Exit Sub
        End If
        If CDate(Me.dtEnd.value) > Date Then
            MsgBox "Date cannot be in the future.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub dtStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.dtStart.value) <> "" Then
        If Not IsDate(Me.dtStart.value) Then
            MsgBox "Please enter a valid date.", vbExclamation
            Cancel = True
            Exit Sub
        End If
        If CDate(Me.dtStart.value) > Date Then
            MsgBox "Date cannot be in the future.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
End Sub


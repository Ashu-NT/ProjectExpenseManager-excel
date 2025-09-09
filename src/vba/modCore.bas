Attribute VB_Name = "modCore"
Option Explicit

' ---------- Helpers ----------
Public Function GetTable(tblName As String) As ListObject
    Dim lo As ListObject, sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = sh.ListObjects(tblName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set GetTable = lo
            Exit Function
        End If
    Next sh
    Set GetTable = Nothing
End Function

Public Function NextID(tblName As String, colName As String) As Long
    Dim lo As ListObject, rng As Range, maxVal As Variant
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        NextID = 1
        Exit Function
    End If
    On Error Resume Next
    Set rng = lo.ListColumns(colName).DataBodyRange
    If rng Is Nothing Then
        NextID = 1
    Else
        maxVal = Application.WorksheetFunction.Max(rng)
        If IsError(maxVal) Or IsEmpty(maxVal) Then
            NextID = 1
        Else
            NextID = CLng(maxVal) + 1
        End If
    End If
    On Error GoTo 0
End Function

' ---------- Audit ----------
Public Sub AuditWrite(action As String, tableName As String, recordID As Variant, userName As String, summary As String)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblAudit")
    If lo Is Nothing Then Exit Sub
    Set lr = lo.ListRows.Add
    lr.Range(lo.ListColumns("AuditID").Index).value = NextID("tblAudit", "AuditID")
    lr.Range(lo.ListColumns("Action").Index).value = action
    lr.Range(lo.ListColumns("TableName").Index).value = tableName
    lr.Range(lo.ListColumns("RecordID").Index).value = recordID
    lr.Range(lo.ListColumns("UserName").Index).value = userName
    lr.Range(lo.ListColumns("TimeStamp").Index).value = Now
    lr.Range(lo.ListColumns("Summary").Index).value = summary
End Sub

' ---------- Lookup helpers ----------
Public Function GetCompanyIDByName(name As String) As Variant
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblCompanies")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.ListColumns("CompanyName").DataBodyRange.rows
        If Trim(r.value) = Trim(name) Then
            GetCompanyIDByName = r.EntireRow.Cells(1, lo.ListColumns("CompanyID").Index).value
            Exit Function
        End If
    Next r
    GetCompanyIDByName = Empty
End Function

Public Function GetCompanyNameByID(compID As Variant) As String
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblCompanies")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.ListColumns("CompanyID").DataBodyRange.rows
        If r.value = compID Then
            GetCompanyNameByID = r.EntireRow.Cells(1, lo.ListColumns("CompanyName").Index).value
            Exit Function
        End If
    Next r
    GetCompanyNameByID = ""
End Function

Public Function GetStatusNameByID(ByVal statusID As Variant) As String
    Dim lo As ListObject, r As Range
    On Error GoTo Done
    Set lo = GetTable("tblLookups")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then GoTo Done
    For Each r In lo.DataBodyRange.rows
        If r.Cells(1, lo.ListColumns("LookupType").Index).value = "ProjectStatus" Then
            If r.Cells(1, lo.ListColumns("LookupID").Index).value = statusID Then
                GetStatusNameByID = r.Cells(1, lo.ListColumns("Value").Index).value
                Exit Function
            End If
        End If
    Next r
Done:
End Function

Public Function GetWorkerIDByName(workerName As String) As Variant
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblWorkers")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.ListColumns("WorkerName").DataBodyRange.rows
        If Trim(r.value) = Trim(workerName) Then
            GetWorkerIDByName = r.EntireRow.Cells(1, lo.ListColumns("WorkerID").Index).value
            Exit Function
        End If
    Next r
    GetWorkerIDByName = Empty
End Function

Public Function GetWorkerNameByID(workerID As Variant) As String
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblWorkers")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.ListColumns("WorkerID").DataBodyRange.rows
        If r.value = workerID Then
            GetWorkerNameByID = r.EntireRow.Cells(1, lo.ListColumns("WorkerName").Index).value
            Exit Function
        End If
    Next r
    GetWorkerNameByID = ""
End Function

' Safe UI helpers ------------------------------------------------
Public Function IsFormLoaded(formName As String) As Boolean
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.name = formName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next uf
    IsFormLoaded = False
End Function

Public Sub RefreshStagingLists_UI()
    On Error GoTo ExitHandler
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.name = "frm_UI" Then
            If TypeOf uf Is Object  Then
                uf.RefreshStagingLists
            End If
            Exit Sub
        End If
    Next uf
ExitHandler:
    ' silently ignore if UI not present or method missing
End Sub

Public Function GetStagingCount() As Long
    Dim cnt As Long
    cnt = 0
    Dim lo As ListObject
    
    Set lo = GetTable("tblStgConsumables")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then cnt = cnt + lo.ListRows.Count
    
    Set lo = GetTable("tblStgPayments")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then cnt = cnt + lo.ListRows.Count
    
    Set lo = GetTable("tblStgLogistics")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then cnt = cnt + lo.ListRows.Count
    
    Set lo = GetTable("tblStgSafety")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then cnt = cnt + lo.ListRows.Count
    
    Set lo = GetTable("tblStgMaterials")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then cnt = cnt + lo.ListRows.Count
    
    GetStagingCount = cnt
End Function

' Small debug helper to show staging counts
Public Sub ShowStagingCounts()
    Dim s As String, lo As ListObject
    Set lo = GetTable("tblStgConsumables")
    s = "Consumables staging: " & IIf(lo Is Nothing Or lo.DataBodyRange Is Nothing, 0, lo.ListRows.Count) & vbCrLf
    Set lo = GetTable("tblStgPayments")
    s = s & "Payments staging: " & IIf(lo Is Nothing Or lo.DataBodyRange Is Nothing, 0, lo.ListRows.Count) & vbCrLf
    Set lo = GetTable("tblStgLogistics")
    s = s & "Logistics staging: " & IIf(lo Is Nothing Or lo.DataBodyRange Is Nothing, 0, lo.ListRows.Count) & vbCrLf
    Set lo = GetTable("tblStgSafety")
    s = s & "Safety staging: " & IIf(lo Is Nothing Or lo.DataBodyRange Is Nothing, 0, lo.ListRows.Count) & vbCrLf
    Set lo = GetTable("tblStgMaterials")
    s = s & "Materials staging: " & IIf(lo Is Nothing Or lo.DataBodyRange Is Nothing, 0, lo.ListRows.Count) & vbCrLf
    MsgBox s, vbInformation, "Staging counts"
End Sub


' ---------- ValidateProjectForm ----------
' Call: If Not ValidateProjectForm(frm_UI) Then Exit Sub
Public Function ValidateProjectForm(frm As Object) As Boolean
    On Error GoTo Fail
    ValidateProjectForm = False

    Dim sName As String, sCode As String, sCompany As String, sStatus As String
    Dim dtStart As String, dtEnd As String
    Dim vBudget As Variant

    sName = Trim$(frm.txtProjectName.value & "")
    sCode = Trim$(frm.txtProjectCode.value & "")
    sCompany = Trim$(frm.cmbCompany.value & "")
    sStatus = Trim$(frm.cmbStatus.value & "")
    dtStart = Trim$(frm.dtStart.value & "")
    dtEnd = Trim$(frm.dtEnd.value & "")
    vBudget = frm.txtBudget.value

    ' 1) Required: Project name
    If sName = "" Then
        MsgBox "Project name is required.", vbExclamation, "Validation"
        frm.txtProjectName.SetFocus
        Exit Function
    End If

    ' 2) Required: Project code
    If sCode = "" Then
        MsgBox "Project code is required.", vbExclamation, "Validation"
        frm.txtProjectCode.SetFocus
        Exit Function
    End If

    ' 3) Company must exist in tblCompanies (use GetCompanyIDByName helper)
    If sCompany = "" Then
        MsgBox "Please select a company (Client).", vbExclamation, "Validation"
        frm.cmbCompany.SetFocus
        Exit Function
    Else
        Dim compID As Variant
        compID = GetCompanyIDByName(sCompany)
        If IsEmpty(compID) Then
            MsgBox "Selected company not found in DB_Companies. Please select an existing company or add it to DB_Companies.", vbExclamation, "Validation"
            frm.cmbCompany.SetFocus
            Exit Function
        End If
    End If

    ' 4) Start date must be a date
    If Not IsDate(dtStart) Then
        MsgBox "Start date is required and must be a valid date.", vbExclamation, "Validation"
        frm.dtStart.SetFocus
        Exit Function
    End If

    ' 5) If End date provided, must be a valid date and >= Start date
    If dtEnd <> "" Then
        If Not IsDate(dtEnd) Then
            MsgBox "End date must be a valid date.", vbExclamation, "Validation"
            frm.dtEnd.SetFocus
            Exit Function
        End If
        If CDate(dtEnd) < CDate(dtStart) Then
            MsgBox "End date cannot be before Start date.", vbExclamation, "Validation"
            frm.dtEnd.SetFocus
            Exit Function
        End If
    End If

    ' 6) Status must be provided and must exist in tblLookups (LookupType = ProjectStatus)
    If sStatus = "" Then
        MsgBox "Please select a project Status.", vbExclamation, "Validation"
        frm.cmbStatus.SetFocus
        Exit Function
    Else
        Dim stID As Variant
        stID = GetLookupIDByTypeAndValue("ProjectStatus", sStatus)
        If IsEmpty(stID) Then
            MsgBox "Selected Status '" & sStatus & "' not found in lookups (tblLookups).", vbExclamation, "Validation"
            frm.cmbStatus.SetFocus
            Exit Function
        End If
    End If

    ' 7) Budget: optional but if provided must be numeric and >= 0
    If Trim$(vBudget & "") <> "" Then
        If Not IsNumeric(vBudget) Then
            MsgBox "Budget must be a numeric value.", vbExclamation, "Validation"
            frm.txtBudget.SetFocus
            Exit Function
        End If
        If val(vBudget) < 0 Then
            MsgBox "Budget cannot be negative.", vbExclamation, "Validation"
            frm.txtBudget.SetFocus
            Exit Function
        End If
    End If

    ' 8) ProjectCode uniqueness (exclude current project when updating)
    If IsProjectCodeDuplicate(sCode, CurrentProjectID) Then
        If MsgBox("Project code '" & sCode & "' already exists. Do you want to continue anyway?", vbYesNo + vbExclamation, "Duplicate Project Code") = vbNo Then
            frm.txtProjectCode.SetFocus
            Exit Function
        End If
    End If

    ' All checks passed
    ValidateProjectForm = True
    Exit Function

Fail:
    MsgBox "Validation error: " & Err.Description, vbCritical, "Validation"
    ValidateProjectForm = False
End Function

' ---------- Helper: check duplicate project code ----------
Private Function IsProjectCodeDuplicate(projCode As String, excludeProjectID As Long) As Boolean
    Dim lo As ListObject, r As Range
    IsProjectCodeDuplicate = False
    If Trim$(projCode) = "" Then Exit Function

    Set lo = GetTable("tblProjects")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    Dim pidCol As Long, codeCol As Long
    pidCol = lo.ListColumns("ProjectID").Index
    codeCol = lo.ListColumns("ProjectCode").Index

    For Each r In lo.DataBodyRange.rows
        If Trim$(CStr(r.Cells(1, codeCol).value)) = Trim$(projCode) Then
            If excludeProjectID > 0 Then
                If CLng(r.Cells(1, pidCol).value) <> CLng(excludeProjectID) Then
                    IsProjectCodeDuplicate = True: Exit Function
                End If
            Else
                IsProjectCodeDuplicate = True: Exit Function
            End If
        End If
    Next r
End Function

' ---------- Helper: get lookup ID by LookupType + Value (returns Empty if not found) ----------
Private Function GetLookupIDByTypeAndValue(lookupType As String, lookupValue As String) As Variant
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblLookups")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    Dim typeCol As Long, valCol As Long, idCol As Long
    typeCol = lo.ListColumns("LookupType").Index
    valCol = lo.ListColumns("Value").Index
    idCol = lo.ListColumns("LookupID").Index

    For Each r In lo.DataBodyRange.rows
        If Trim$(CStr(r.Cells(1, typeCol).value)) = Trim$(lookupType) _
           And Trim$(CStr(r.Cells(1, valCol).value)) = Trim$(lookupValue) Then
            GetLookupIDByTypeAndValue = r.Cells(1, idCol).value
            Exit Function
        End If
    Next r
End Function




' Return projectID if ProjectCode exists (case-insensitive), else 0
Public Function FindProjectByCode(projectCode As String) As Long
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblProjects")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.ListColumns("ProjectCode").DataBodyRange.rows
        If LCase(Trim(r.value)) = LCase(Trim(projectCode)) Then
            FindProjectByCode = r.EntireRow.Cells(1, lo.ListColumns("ProjectID").Index).value
            Exit Function
        End If
    Next r
    FindProjectByCode = 0
End Function

' Return True if any staging rows exist
Public Function HasAnyStagingRows() As Boolean
    Dim lo As ListObject
    Set lo = GetTable("tblStgConsumables")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then HasAnyStagingRows = True
    Set lo = GetTable("tblStgPayments")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then HasAnyStagingRows = True
    Set lo = GetTable("tblStgLogistics")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then HasAnyStagingRows = True
    Set lo = GetTable("tblStgSafety")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then HasAnyStagingRows = True
    Set lo = GetTable("tblStgMaterials")
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then HasAnyStagingRows = True
End Function

Public Function CreateProjectFromForm(frm As Object) As Long
    Dim loProj As ListObject, lr As ListRow, newID As Long
    Set loProj = GetTable("tblProjects")
    If loProj Is Nothing Then MsgBox "tblProjects missing": Exit Function
    ' Validation already done in caller
    newID = NextID("tblProjects", "ProjectID")
    Set lr = loProj.ListRows.Add
    lr.Range(loProj.ListColumns("ProjectID").Index).value = newID
    lr.Range(loProj.ListColumns("ProjectCode").Index).value = Trim(frm.txtProjectCode.value)
    lr.Range(loProj.ListColumns("ProjectName").Index).value = Trim(frm.txtProjectName.value)
    lr.Range(loProj.ListColumns("CompanyID").Index).value = GetCompanyIDByName(Trim(frm.cmbCompany.value))
    lr.Range(loProj.ListColumns("StartDate").Index).value = CDate(frm.dtStart.value)
    If Trim(frm.dtEnd.value) <> "" Then lr.Range(loProj.ListColumns("EndDate").Index).value = CDate(frm.dtEnd.value)
    lr.Range(loProj.ListColumns("Budget").Index).value = CDbl(frm.txtBudget.value)
    lr.Range(loProj.ListColumns("ProjectManager").Index).value = Trim(frm.txtManager.value)
    lr.Range(loProj.ListColumns("Status").Index).value = Trim(frm.cmbStatus.value)
    lr.Range(loProj.ListColumns("Notes").Index).value = Trim(frm.txtNotes.value)
    AuditWrite "Create", "tblProjects", newID, Environ$("USERNAME"), "Project created via UI"
    CurrentProjectID = newID
    CreateProjectFromForm = newID
End Function




' ---- GetProjectsForSearch (returns 2D array with 5 columns) ----
' Columns returned (in order):
'   0 = ProjectID (Long)
'   1 = ProjectCode (String)
'   2 = ProjectName (String)
'   3 = StartDate (String, formatted yyyy-mm-dd or blank)
'   4 = Status (String)
'
' If term = "" then returns ALL projects.
' Returns a zero-length Variant array if no rows found.
Public Function GetProjectsForSearch(ByVal term As String) As Variant
    Dim lo As ListObject
    Dim r As Range
    Dim idCol As Long, codeCol As Long, nameCol As Long, startCol As Long, statusCol As Long
    Dim tempList As Collection
    Dim rec As Variant
    Dim v() As Variant
    Dim i As Long, cnt As Long
    Dim sTerm As String
    
    term = Trim$(term)
    sTerm = LCase$(term)
    
    Set lo = GetTable("tblProjects")
    If lo Is Nothing Then
        GetProjectsForSearch = Array()   ' empty
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        GetProjectsForSearch = Array()   ' empty
        Exit Function
    End If
    
    ' Try to get column indexes; if missing, set to 0 and skip values later
    On Error Resume Next
    idCol = lo.ListColumns("ProjectID").Index: If Err.Number <> 0 Then idCol = 0: Err.Clear
    codeCol = lo.ListColumns("ProjectCode").Index: If Err.Number <> 0 Then codeCol = 0: Err.Clear
    nameCol = lo.ListColumns("ProjectName").Index: If Err.Number <> 0 Then nameCol = 0: Err.Clear
    startCol = lo.ListColumns("StartDate").Index: If Err.Number <> 0 Then startCol = 0: Err.Clear
    statusCol = lo.ListColumns("Status").Index: If Err.Number <> 0 Then statusCol = 0: Err.Clear
    On Error GoTo 0
    
    Set tempList = New Collection
    cnt = 0
    For Each r In lo.DataBodyRange.rows
        Dim projID As Variant, projCode As String, projName As String, projStart As String, projStatus As String
        projID = IIf(idCol > 0, r.Cells(1, idCol).value, Empty)
        projCode = IIf(codeCol > 0, CStr(r.Cells(1, codeCol).value), "")
        projName = IIf(nameCol > 0, CStr(r.Cells(1, nameCol).value), "")
        If startCol > 0 Then
            If IsDate(r.Cells(1, startCol).value) Then
                projStart = Format$(CDate(r.Cells(1, startCol).value), "yyyy-mm-dd")
            Else
                projStart = CStr(r.Cells(1, startCol).value)
            End If
        Else
            projStart = ""
        End If
        projStatus = IIf(statusCol > 0, CStr(r.Cells(1, statusCol).value), "")
        
        ' If term empty, include all; otherwise include if name or code contains term (case-insensitive)
        If sTerm = "" Then
            rec = Array(projID, projCode, projName, projStart, projStatus)
            tempList.Add rec
            cnt = cnt + 1
        Else
            If (InStr(1, LCase$(projName), sTerm, vbBinaryCompare) > 0) _
            Or (InStr(1, LCase$(projCode), sTerm, vbBinaryCompare) > 0) Then
                rec = Array(projID, projCode, projName, projStart, projStatus)
                tempList.Add rec
                cnt = cnt + 1
            End If
        End If
    Next r
    
    If cnt = 0 Then
        GetProjectsForSearch = Array()   ' empty
        Exit Function
    End If
    
    ' Build the 2D array (0..cnt-1, 0..4)
    ReDim v(0 To cnt - 1, 0 To 4)
    For i = 1 To cnt
        v(i - 1, 0) = tempList(i)(0)
        v(i - 1, 1) = tempList(i)(1)
        v(i - 1, 2) = tempList(i)(2)
        v(i - 1, 3) = tempList(i)(3)
        v(i - 1, 4) = tempList(i)(4)
    Next i
    
    GetProjectsForSearch = v
End Function


' Deletes all data rows from the specified table (keeps header row)
Public Sub ClearTable(lo As ListObject)
    On Error Resume Next
    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    On Error GoTo 0
End Sub

' Copies rows from a source (DB) table with ProjectID match into a staging table.
' fieldList is a simple VBA Array of column-names (strings) present in both tables.
Public Sub CopyProjectRowsToStaging(ByVal srcName As String, ByVal stgName As String, _
                                    ByVal projectID As Long, ByVal fieldList As Variant)
    On Error GoTo ErrHandler
    Dim loSrc As ListObject, loStg As ListObject, r As Range, lr As ListRow
    Dim i As Long, srcProjCol As Long
    Dim fld As Variant, srcColIndex As Long, stgColIndex As Long

    Set loSrc = GetTable(srcName)
    Set loStg = GetTable(stgName)
    If loSrc Is Nothing Then Err.Raise vbObjectError + 1, , "Source table '" & srcName & "' not found."
    If loStg Is Nothing Then Err.Raise vbObjectError + 2, , "Staging table '" & stgName & "' not found."
    If loSrc.DataBodyRange Is Nothing Then Exit Sub

    ' Make sure ProjectID exists in source
    If Not ColumnExists(loSrc, "ProjectID") Then
        Err.Raise vbObjectError + 3, , "Source table '" & srcName & "' does not have ProjectID."
    End If
    srcProjCol = loSrc.ListColumns("ProjectID").Index

    For Each r In loSrc.DataBodyRange.rows
        If r.Cells(1, srcProjCol).value = projectID Then
            Set lr = loStg.ListRows.Add
            ' Set a new TempID if the staging has one
            If ColumnExists(loStg, "TempID") Then
                lr.Range.Cells(1, loStg.ListColumns("TempID").Index).value = NextID(stgName, "TempID")
            End If

            ' copy fields in order (only if both src and stg columns exist)
            For i = LBound(fieldList) To UBound(fieldList)
                fld = fieldList(i)
                ' get source column index if exists
                If ColumnExists(loSrc, fld) Then
                    srcColIndex = loSrc.ListColumns(CStr(fld)).Index
                Else
                    srcColIndex = 0
                End If
                ' get staging column index if exists
                If ColumnExists(loStg, fld) Then
                    stgColIndex = loStg.ListColumns(CStr(fld)).Index
                Else
                    stgColIndex = 0
                End If

                If srcColIndex > 0 And stgColIndex > 0 Then
                    ' lr.Range is the row Range for the new ListRow.
                    ' Use Cells(1, stgColIndex) where stgColIndex is table column index.
                    lr.Range.Cells(1, stgColIndex).value = r.Cells(1, srcColIndex).value
                End If
            Next i
        End If
    Next r

    Exit Sub

ErrHandler:
    MsgBox "Error in CopyProjectRowsToStaging(" & srcName & "," & stgName & "): " & Err.Number & " - " & Err.Description, vbExclamation
End Sub


' Returns True if a named column exists in the given ListObject.
' Accepts flexible types to avoid ByRef type-mismatch issues.
Public Function ColumnExists(lo As ListObject, colName As Variant) As Boolean
    On Error GoTo HandleErr
    If lo Is Nothing Then
        ColumnExists = False
        Exit Function
    End If
    ' Force colName to string (safe)
    Dim sCol As String
    sCol = CStr(colName)
    Dim testCol As ListColumn
    Set testCol = lo.ListColumns(sCol)
    ColumnExists = Not testCol Is Nothing
    Exit Function

HandleErr:
    ColumnExists = False
    Err.Clear
End Function


' ---------- resolve lookup----------

Public Function GetLookupID(lookupType As String, lookupValue As String) As Variant
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblLookups")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.DataBodyRange.rows
        If Trim(r.Cells(lo.ListColumns("LookupType").Index).value) = lookupType _
           And Trim(r.Cells(lo.ListColumns("Value").Index).value) = Trim(lookupValue) Then
            GetLookupID = r.Cells(lo.ListColumns("LookupID").Index).value
            Exit Function
        End If
    Next r
    GetLookupID = Empty
End Function


' --- staging helpers ---
Public Function CountStagingRows(tblName As String) As Long
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        CountStagingRows = 0
        Exit Function
    End If
    On Error Resume Next
    If lo.DataBodyRange Is Nothing Then
        CountStagingRows = 0
    Else
        CountStagingRows = lo.ListRows.Count
    End If
    On Error GoTo 0
End Function



' Confirm commit: returns True if user wants to proceed committing staging rows
Public Function ConfirmStagingCommit(Optional warnAboutDuplicatesIfLoaded As Boolean = True) As Boolean
    Dim cCons As Long, cPay As Long, cLog As Long, cSafe As Long, cMat As Long
    cCons = CountStagingRows("tblStgConsumables")
    cPay = CountStagingRows("tblStgPayments")
    cLog = CountStagingRows("tblStgLogistics")
    cSafe = CountStagingRows("tblStgSafety")
    cMat = CountStagingRows("tblStgMaterials")
    
    Dim total As Long: total = cCons + cPay + cLog + cSafe + cMat
    If total = 0 Then
        ConfirmStagingCommit = False
        Exit Function
    End If

    Dim msg As String
    msg = "Staging contains:" & vbCrLf & _
          "  Consumables: " & cCons & vbCrLf & _
          "  Payments:    " & cPay & vbCrLf & _
          "  Logistics:    " & cLog & vbCrLf & _
          "  Safety Item:    " & cSafe & vbCrLf & _
          "  Materials:   " & cMat & vbCrLf & vbCrLf

    If warnAboutDuplicatesIfLoaded And CurrentProjectID > 0 Then
        msg = msg & "WARNING: You have a project currently loaded (ProjectID = " & CurrentProjectID & ")." & vbCrLf & _
                    "If these rows were copied from the database when the project was loaded, committing them now will INSERT them again and create duplicates." & vbCrLf & vbCrLf & _
                    "If you intend to only add NEW rows, click NO and then use the 'Clear Loaded Items' button to clear the staging tables before adding new lines." & vbCrLf & vbCrLf & _
                    "Click YES to proceed and commit staging rows to the current project (duplicates may result)."
    Else
        msg = msg & "Do you want to commit these staging rows to the project?"
    End If

    ConfirmStagingCommit = (MsgBox(msg, vbYesNo + vbExclamation, "Confirm commit staging rows") = vbYes)
End Function




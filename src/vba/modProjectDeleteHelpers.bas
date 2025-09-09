Attribute VB_Name = "modProjectDeleteHelpers"
Option Explicit

' --------- Permission check for deletion ----------
Public Function CanUserDeleteProject() As Boolean
    Dim allowed As String, curUser As String
    curUser = Environ$("USERNAME")
    allowed = GetSetting("AllowProjectDeleteUsernames", "") ' comma-separated usernames (optional)

    If Trim(allowed) <> "" Then
        If IsUserInCsv(allowed, curUser) Then
            CanUserDeleteProject = True
            Exit Function
        End If
    End If

    ' If not in allowed list, require Admin password
    Dim pwd As String
    pwd = InputBox("Enter Admin password to authorize project deletion:", "Admin authentication")
    If StrPtr(pwd) = 0 Then Exit Function
    If VerifyAdminPassword(pwd) Then CanUserDeleteProject = True
End Function

' --------- make a backup copy of the workbook (simple, saves to Backups folder) ----------
Public Sub CreateBackupCopy(Optional tag As String = "")
    On Error GoTo ErrHandler
    Dim wb As Workbook, fpath As String, fname As String, bkdir As String, stamp As String
    Set wb = ThisWorkbook
    fpath = wb.Path
    If fpath = "" Then
        ' workbook not saved yet — save a temporary copy to Temp folder
        fpath = Environ$("TEMP")
    End If
    bkdir = fpath & Application.PathSeparator & "Backups"
    If Dir(bkdir, vbDirectory) = "" Then MkDir bkdir
    stamp = Format(Now, "yyyy-mm-dd_HHMMSS")
    fname = wb.name
    ' remove extension and append tag + timestamp
    If InStrRev(fname, ".") > 0 Then fname = Left(fname, InStrRev(fname, ".") - 1)
    Dim outName As String
    If Trim(tag) <> "" Then
        outName = fname & "_" & tag & "_" & stamp & ".xlsm"
    Else
        outName = fname & "_backup_" & stamp & ".xlsm"
    End If
    wb.SaveCopyAs bkdir & Application.PathSeparator & outName
    Exit Sub
ErrHandler:
    ' ignore backup errors but inform developer in Immediate
    Debug.Print "CreateBackupCopy error: " & Err.Description
    Err.Clear
End Sub

' --------- Delete project + all linked DB rows and staging rows ----------
Public Function DeleteProjectByID(ByVal projectID As Long) As Boolean
    On Error GoTo ErrHandler
    Dim lo As ListObject
    Dim tblNames As Variant
    Dim i As Long, colPID As Long

    If projectID <= 0 Then Exit Function

    ' First: delete rows from child DB tables
    tblNames = Array("tblConsumables", "tblPayments", "tblLogistics", "tblSafety", "tblMaterials")
    For i = LBound(tblNames) To UBound(tblNames)
        Set lo = Nothing
        Set lo = GetTable(CStr(tblNames(i)))
        If Not lo Is Nothing Then
            colPID = ColIndex(lo, "ProjectID")
            If colPID > 0 Then
                Dim rCount As Long
                ' iterate bottom->top deleting matching rows
                For rCount = lo.ListRows.Count To 1 Step -1
                    If lo.DataBodyRange.rows(rCount).Cells(colPID).value = projectID Then
                        lo.ListRows(rCount).Delete
                    End If
                Next rCount
            End If
        End If
    Next i

    ' Next: delete staging rows that may reference the project (if you stored ProjectID in staging)
    tblNames = Array("tblStgConsumables", "tblStgPayments", "tblStgLogistics", "tblStgSafety", "tblStgMaterials")
    For i = LBound(tblNames) To UBound(tblNames)
        Set lo = Nothing
        Set lo = GetTable(CStr(tblNames(i)))
        If Not lo Is Nothing Then
            colPID = ColIndex(lo, "ProjectID") ' some staging may not have ProjectID; safe return 0 if missing
            If colPID > 0 Then
                Dim rCount2 As Long
                For rCount2 = lo.ListRows.Count To 1 Step -1
                    If lo.DataBodyRange.rows(rCount2).Cells(colPID).value = projectID Then
                        lo.ListRows(rCount2).Delete
                    End If
                Next rCount2
            Else
                ' If no ProjectID column exists in staging, we still clear all staging rows if they belong to the loaded project context
                ' (optional) skip to avoid accidental deletes
            End If
        End If
    Next i

    ' Finally: delete the project row from tblProjects
    Set lo = GetTable("tblProjects")
    If Not lo Is Nothing Then
        Dim idx As Long, found As Boolean
        found = False
        colPID = ColIndex(lo, "ProjectID")
        If colPID > 0 Then
            For idx = lo.ListRows.Count To 1 Step -1
                If lo.DataBodyRange.rows(idx).Cells(colPID).value = projectID Then
                    lo.ListRows(idx).Delete
                    found = True
                    Exit For
                End If
            Next idx
        End If
        If Not found Then
            ' no matching project found — return false
            DeleteProjectByID = False
            Exit Function
        End If
    Else
        DeleteProjectByID = False
        Exit Function
    End If

    ' Audit: write a deletion record (if you have AuditWrite)
    On Error Resume Next
    AuditWrite "Delete", "tblProjects", projectID, Environ$("USERNAME"), "Project and child rows deleted"
    On Error GoTo ErrHandler

    DeleteProjectByID = True
    Exit Function

ErrHandler:
    Debug.Print "DeleteProjectByID error: " & Err.Description
    DeleteProjectByID = False
    Err.Clear
End Function

' --------- Clear the form project fields (adjust to your controls) ----------
Public Sub ClearProjectForm(frm As Object)
    On Error Resume Next
    frm.txtProjectName.value = ""
    frm.txtProjectCode.value = ""
    frm.cmbCompany.value = ""
    frm.dtStart.value = ""
    frm.dtEnd.value = ""
    frm.txtBudget.value = ""
    frm.txtManager.value = ""
    frm.cmbStatus.value = ""
    frm.txtNotes.value = ""
    ' Clear staging lists if desired
    ClearTable GetTable("tblStgConsumables")
    ClearTable GetTable("tblStgPayments")
    ClearTable GetTable("tblStgLogistics")
    On Error GoTo 0
End Sub


Attribute VB_Name = "modExportAll"
Option Explicit

' ExportAllVBA
' Exports all VBA components from the current workbook into a folder: <ThisWorkbook.Path>\src\vba
' - Requires: "Trust access to the VBA project object model" enabled in Excel Trust Center.
' - Will fail if the VBProject is password-protected.
Public Sub ExportAllVBA(Optional ByVal targetFolder As String = "")
    On Error GoTo ErrHandler
    Dim vbProj As Object
    Dim vbComp As Object
    Dim exportPath As String
    Dim exportedCount As Long

    ' default export location: <workbook folder>\src\vba
    If Trim(targetFolder) = "" Then
        exportPath = ThisWorkbook.Path & "\src\vba"
    Else
        exportPath = targetFolder
    End If

    ' Quick check: can we access the VBProject?
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        MsgBox "Cannot access the VBA project. Please enable 'Trust access to the VBA project object model' in Excel Trust Center and try again.", vbExclamation, "ExportAllVBA"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' Ensure base folders exist (simple creation under ThisWorkbook.Path)
    Call EnsureExportFolderExists(exportPath)

    exportedCount = 0
    For Each vbComp In vbProj.VBComponents
        Dim compType As Long
        Dim fname As String
        compType = vbComp.Type

        Select Case compType
            Case 1 ' vbext_ct_StdModule
                fname = exportPath & "\" & vbComp.name & ".bas"
            Case 2 ' vbext_ct_ClassModule
                fname = exportPath & "\" & vbComp.name & ".cls"
            Case 3 ' vbext_ct_MSForm (UserForm)
                fname = exportPath & "\" & vbComp.name & ".frm"
            Case 100 ' vbext_ct_Document (ThisWorkbook / Worksheets)
                fname = exportPath & "\" & vbComp.name & ".cls"
            Case Else
                fname = exportPath & "\" & vbComp.name & ".txt"
        End Select

        ' If a file exists, attempt to delete first (overwrite)
        If Dir(fname) <> "" Then
            On Error Resume Next
            Kill fname
            On Error GoTo ErrHandler
        End If

        ' Export component - will create .frm and .frx for forms automatically
        vbComp.Export fname
        exportedCount = exportedCount + 1
        Debug.Print "Exported: " & fname
    Next vbComp

    MsgBox "Export complete: " & exportedCount & " components exported to:" & vbCrLf & exportPath, vbInformation, "ExportAllVBA"
    Exit Sub

ErrHandler:
    MsgBox "ExportAllVBA error: " & Err.Number & " - " & Err.Description, vbCritical, "ExportAllVBA"
End Sub


' EnsureExportFolderExists: creates folder(s) if missing.
' This routine handles the common default: <workbook folder>\src\vba
Private Sub EnsureExportFolderExists(ByVal fullPath As String)
    Dim basePath As String, srcPath As String, vbaPath As String

    ' Normalize: remove trailing backslash
    If Right(fullPath, 1) = "\" Then fullPath = Left(fullPath, Len(fullPath) - 1)

    ' If user passed a path under ThisWorkbook.Path, create nested folders safely.
    If InStr(1, fullPath, ThisWorkbook.Path, vbTextCompare) = 1 Then
        ' expected default like C:\...\MyWorkbook\src\vba
        ' create progressively
        Dim parts() As String, i As Long, accum As String
        parts = Split(mID(fullPath, Len(ThisWorkbook.Path) + 2), "\") ' parts after workbook path
        accum = ThisWorkbook.Path
        For i = LBound(parts) To UBound(parts)
            accum = accum & "\" & parts(i)
            If Dir(accum, vbDirectory) = "" Then
                MkDir accum
            End If
        Next i
    Else
        ' If user passed arbitrary folder, try to create that full path if immediate parent exists.
        If Dir(fullPath, vbDirectory) = "" Then
            On Error Resume Next
            MkDir fullPath
            If Err.Number <> 0 Then
                ' Try create parent then this folder (one level)
                Err.Clear
                Dim parentPath As String
                parentPath = Left(fullPath, InStrRev(fullPath, "\") - 1)
                If Dir(parentPath, vbDirectory) = "" Then
                    ' cannot create because parent doesn't exist - give informative error
                    MsgBox "Cannot create folder '" & fullPath & "' because parent '" & parentPath & "' does not exist. Use a folder under the workbook path or create parent folders manually.", vbExclamation, "EnsureExportFolderExists"
                    Exit Sub
                Else
                    MkDir fullPath
                End If
            End If
            On Error GoTo 0
        End If
    End If
End Sub


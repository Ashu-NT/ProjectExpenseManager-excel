Attribute VB_Name = "modDebug"
Public Sub ShowSheetVisibilityReport()
    Dim ws As Worksheet, s As String
    s = "Sheet visibility report:" & vbCrLf
    For Each ws In ThisWorkbook.Worksheets
        s = s & ws.name & " -> Visible=" & ws.Visible & ", Protected=" & ws.ProtectContents & vbCrLf
    Next ws
    Debug.Print s
    MsgBox "Sheet visibility printed to Immediate (Ctrl+G).", vbInformation
End Sub

Public Sub ForceUnprotectSettingsNow()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    If ws Is Nothing Then
        MsgBox "Settings sheet not found.", vbExclamation
        Exit Sub
    End If

    If TryUnprotectSheet(ws) Then
        MsgBox "Settings sheet unprotected (one of candidate passwords worked).", vbInformation
    Else
        MsgBox "Could not unprotect Settings sheet automatically. Next steps:" & vbCrLf & _
               "- Try AdminUnlockAllSheets (enter Admin password) or" & vbCrLf & _
               "- Manually unprotect the sheet using the password used when it was first created.", vbExclamation
    End If
End Sub


Public Sub SetSetting(ByVal key As String, ByVal value As String)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblSettings")
    If lo Is Nothing Then Exit Sub

    On Error Resume Next
    Set lr = Nothing
    Set lr = lo.ListColumns("Key").DataBodyRange.Find(What:=key, LookAt:=xlWhole)
    On Error GoTo 0

    If lr Is Nothing Then
        Set lr = lo.ListRows.Add
        lr.Range(lo.ListColumns("Key").Index).value = key
    End If
    lr.EntireRow.Cells(1, lo.ListColumns("Value").Index).value = value
End Sub

Public Sub DumpFormControls_Safety()
    Dim f As Object, c As Control
    Set f = New frm_SafetyLine
    Debug.Print "Controls on frm_SafetyLine:"
    For Each c In f.Controls
        Debug.Print "  " & c.name & " (" & TypeName(c) & ")  Visible=" & c.Visible & "  Enabled=" & c.Enabled
    Next c
    Unload f
End Sub

Public Sub DumpFormControls_Materials()
    Dim f As Object, c As Control
    Set f = New frm_MaterialLIne
    Debug.Print "Controls on frm_MaterialLine:"
    For Each c In f.Controls
        Debug.Print "  " & c.name & " (" & TypeName(c) & ")  Visible=" & c.Visible & "  Enabled=" & c.Enabled
    Next c
    Unload f
End Sub


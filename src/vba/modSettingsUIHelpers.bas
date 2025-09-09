Attribute VB_Name = "modSettingsUIHelpers"
Option Explicit
' ---------- modSettingsUIHelpers ----------
' Helpers to read/write tblSettings with Notes support.


' Get setting value (existing function may do similar) - returns "" if missing
Public Function GetSettingValue(key As String, Optional defaultValue As Variant) As Variant
    Dim lo As ListObject, r As Range
    Set lo = GetSettingsTable
    If lo Is Nothing Then
        GetSettingValue = IIf(IsMissing(defaultValue), vbNullString, defaultValue)
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        GetSettingValue = IIf(IsMissing(defaultValue), vbNullString, defaultValue)
        Exit Function
    End If
    For Each r In lo.ListColumns("Key").DataBodyRange.rows
        If Trim(CStr(r.value)) = key Then
            GetSettingValue = CStr(r.Offset(0, 1).value)
            Exit Function
        End If
    Next r
    GetSettingValue = IIf(IsMissing(defaultValue), vbNullString, defaultValue)
End Function

' Get Notes for a given key (returns "" if none or missing)
Public Function GetSettingNotes(key As String) As String
    Dim lo As ListObject, r As Range, notesCol As Long
    Set lo = GetSettingsTable
    If lo Is Nothing Then Exit Function
    notesCol = ColIndex(lo, "Notes")
    If notesCol = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.ListColumns("Key").DataBodyRange.rows
        If Trim(CStr(r.value)) = key Then
            GetSettingNotes = CStr(r.Offset(0, notesCol - lo.ListColumns("Key").Index).value)
            Exit Function
        End If
    Next r
End Function

' Add or update a setting with notes
Public Sub UpsertSetting(key As String, value As String, Optional notes As String = "")
    Dim lo As ListObject, r As Range, added As Boolean, notesCol As Long
    Set lo = GetSettingsTable
    If lo Is Nothing Then
        MsgBox "Settings table 'tblSettings' not found.", vbCritical
        Exit Sub
    End If

    ' Ensure notes column exists
    notesCol = ColIndex(lo, "Notes")
    If notesCol = 0 Then
        lo.ListColumns.Add.name = "Notes"
        notesCol = ColIndex(lo, "Notes")
    End If

    added = False
    If Not lo.DataBodyRange Is Nothing Then
        For Each r In lo.ListColumns("Key").DataBodyRange.rows
            If Trim(CStr(r.value)) = key Then
                r.Offset(0, 1).value = CStr(value)             ' Value column
                r.Offset(0, notesCol - lo.ListColumns("Key").Index).value = CStr(notes)
                added = True
                Exit For
            End If
        Next r
    End If

    If Not added Then
        Dim lr As ListRow
        Set lr = lo.ListRows.Add
        lr.Range(lo.ListColumns("Key").Index).value = key
        lr.Range(lo.ListColumns("Value").Index).value = value
        lr.Range(lo.ListColumns("Notes").Index).value = notes
    End If

    ' Audit write if available
    On Error Resume Next
    AuditWrite "UpsertSetting", "tblSettings", key, Environ$("USERNAME"), "Value: " & value & " ; Notes: " & Left(notes, 200)
    On Error GoTo 0
End Sub

' Delete a setting by key
Public Function DeleteSettingByKey(key As String) As Boolean
    Dim lo As ListObject, r As Range
    Set lo = GetSettingsTable
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.ListColumns("Key").DataBodyRange.rows
        If Trim(CStr(r.value)) = key Then
            r.EntireRow.Delete
            DeleteSettingByKey = True
            On Error Resume Next
            AuditWrite "DeleteSetting", "tblSettings", key, Environ$("USERNAME"), "Deleted setting"
            On Error GoTo 0
            Exit Function
        End If
    Next r
End Function

' Return all keys as array (useful)
Public Function AllSettingKeys() As Variant
    Dim lo As ListObject, r As Range, arr() As String, i As Long
    Set lo = GetSettingsTable
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        AllSettingKeys = Array()
        Exit Function
    End If
    ReDim arr(0 To lo.ListRows.Count - 1)
    i = 0
    For Each r In lo.ListColumns("Key").DataBodyRange.rows
        arr(i) = CStr(r.value)
        i = i + 1
    Next r
    AllSettingKeys = arr
End Function

' Validation rule: protect certain keys from deletion/edit (case-sensitive list)
Public Function IsProtectedSetting(key As String) As Boolean
    Dim protectedKeys As Variant
    protectedKeys = Array("AdminPassword_Obf", "FormAccessPassword_Obf", "AdminPassword_Obf") ' add more if needed
    Dim k As Variant
    For Each k In protectedKeys
        If StrComp(k, key, vbBinaryCompare) = 0 Then
            IsProtectedSetting = True
            Exit Function
        End If
    Next k
End Function
' ---------- end modSettingsUIHelpers ----------


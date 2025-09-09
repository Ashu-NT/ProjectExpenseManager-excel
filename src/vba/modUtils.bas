Attribute VB_Name = "modUtils"

Option Explicit

' Returns True if the given numeric id exists in the TempID column of the staging table.
Public Function IsIDInStaging(ByVal stagingTableName As String, ByVal testID As Long) As Boolean
    Dim lo As ListObject, rng As Range, r As Range
    On Error Resume Next
    Set lo = GetTable(stagingTableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ' Make sure the TempID column exists
    Dim ci As Long
    ci = ColIndex(lo, "TempID")
    If ci = 0 Then Exit Function
    For Each r In lo.ListColumns("TempID").DataBodyRange.rows
        If val(r.value) = testID Then
            IsIDInStaging = True
            Exit Function
        End If
    Next r
End Function


' Return numeric ID from a ListBox selection. Returns 0 if none or non-numeric.
Public Function GetSelectedIDFromListbox(lst As MSForms.ListBox) As Long
    Dim idx As Long
    If lst.ListCount = 0 Then Exit Function
    idx = lst.ListIndex
    If idx < 0 Then Exit Function
    On Error Resume Next
    ' If multi-column, use column 0; if single column with "ID | Desc", try parse
    If lst.ColumnCount > 1 Then
        GetSelectedIDFromListbox = CLng(lst.List(idx, 0))
    Else
        Dim s As String
        s = CStr(lst.List(idx))
        ' Try parse if "ID | Desc"
        If InStr(s, "|") > 0 Then
            s = Trim(Split(s, "|")(0))
            GetSelectedIDFromListbox = CLng(s)
        Else
            ' Single value – try numeric
            If IsNumeric(s) Then GetSelectedIDFromListbox = CLng(s)
        End If
    End If
    On Error GoTo 0
End Function


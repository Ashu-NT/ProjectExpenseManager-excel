Attribute VB_Name = "modListAllTables"
Public Sub ListAllTables()
    Dim ws As Worksheet, lo As ListObject, out As String
    out = ""
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            out = out & "Table: [" & lo.name & "] on sheet: " & ws.name & vbCrLf
        Next lo
    Next ws
    MsgBox out, vbInformation, "List of tables"
End Sub


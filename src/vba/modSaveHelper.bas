Attribute VB_Name = "modSaveHelper"
Option Explicit

Public Function CheckMaxRowsBeforeCommit() As Boolean
    Dim maxRows As Long
    maxRows = CLng(val(GetSetting("MaxRowsPerCommit", "500")))
    Dim totalRows As Long
    If Not GetTable("tblStgConsumables") Is Nothing Then If Not GetTable("tblStgConsumables").DataBodyRange Is Nothing Then totalRows = totalRows + GetTable("tblStgConsumables").ListRows.Count
    If Not GetTable("tblStgPayments") Is Nothing Then If Not GetTable("tblStgPayments").DataBodyRange Is Nothing Then totalRows = totalRows + GetTable("tblStgPayments").ListRows.Count
    If Not GetTable("tblStgLogistics") Is Nothing Then If Not GetTable("tblStgLogistics").DataBodyRange Is Nothing Then totalRows = totalRows + GetTable("tblStgLogistics").ListRows.Count
    If Not GetTable("tblStgSafety") Is Nothing Then If Not GetTable("tblStgSafety").DataBodyRange Is Nothing Then totalRows = totalRows + GetTable("tblStgSafety").ListRows.Count
    If Not GetTable("tblStgMaterials") Is Nothing Then If Not GetTable("tblStgMaterials").DataBodyRange Is Nothing Then totalRows = totalRows + GetTable("tblStgMaterials").ListRows.Count

    If totalRows = 0 Then CheckMaxRowsBeforeCommit = True: Exit Function
    If totalRows > maxRows Then
        If MsgBox("Staging contains " & totalRows & " rows which exceeds MaxRowsPerCommit (" & maxRows & "). Proceed anyway?", vbYesNo + vbExclamation, "Too many rows") = vbNo Then
            CheckMaxRowsBeforeCommit = False
            Exit Function
        End If
    End If
    CheckMaxRowsBeforeCommit = True
End Function


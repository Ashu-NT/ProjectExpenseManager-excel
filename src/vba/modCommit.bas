Attribute VB_Name = "modCommit"
Option Explicit

' Commits staging rows into DB tables. Returns a small summary string.
Public Function CommitStagingToDB(ByVal projectID As Long, ByVal userName As String) As String
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long, c5 As Long
    c1 = CommitStgConsumables(projectID, userName)
    c2 = CommitStgPayments(projectID, userName)
    c3 = CommitStgLogistics(projectID, userName)
    c4 = CommitStgSafety(projectID, userName)
    c5 = CommitStgMaterials(projectID, userName)
    CommitStagingToDB = "Committed: " & c1 & " consumable(s), " & c2 & " payment(s), " & c3 & " logistic(s) " & c4 & " Safety Item(s) " & c5 & " Material(s)."
End Function

' ---------- Consumables ----------
Public Function CommitStgConsumables(ByVal projectID As Long, ByVal userName As String) As Long
    Dim stg As ListObject, db As ListObject
    Dim i As Long, newRow As ListRow, srcR As Range
    Dim ciDate As Long, ciCategory As Long, ciItem As Long, ciQty As Long, ciUnit As Long, ciSupplier As Long
    Dim dbCID As Long, dbPID As Long, dbDate As Long

    Set stg = GetTable("tblStgConsumables")
    Set db = GetTable("tblConsumables")
    If stg Is Nothing Or db Is Nothing Then Exit Function

    ' Cache column indexes (0 means missing)
    ciDate = ColIndex(stg, "Date")
    ciCategory = ColIndex(stg, "CategoryID")
    ciItem = ColIndex(stg, "ItemDescription")
    ciQty = ColIndex(stg, "Quantity")
    ciUnit = ColIndex(stg, "UnitCost")
    ciSupplier = ColIndex(stg, "Supplier")

    dbCID = ColIndex(db, "ConsumableID")
    dbPID = ColIndex(db, "ProjectID")
    dbDate = ColIndex(db, "Date")

    On Error GoTo ErrHandler
    For i = stg.ListRows.Count To 1 Step -1
        Set srcR = stg.DataBodyRange.rows(i)

        ' Add row to DB
        Set newRow = db.ListRows.Add
        If dbCID > 0 Then newRow.Range.Cells(1, dbCID).value = NextID("tblConsumables", "ConsumableID")
        If dbPID > 0 Then newRow.Range.Cells(1, dbPID).value = projectID
        If dbDate > 0 And ciDate > 0 Then newRow.Range.Cells(1, dbDate).value = srcR.Cells(1, ciDate).value

        ' Safe writes for other fields if columns exist
        If ColIndex(db, "CategoryID") > 0 And ciCategory > 0 Then newRow.Range.Cells(1, ColIndex(db, "CategoryID")).value = srcR.Cells(1, ciCategory).value
        If ColIndex(db, "ItemDescription") > 0 And ciItem > 0 Then newRow.Range.Cells(1, ColIndex(db, "ItemDescription")).value = srcR.Cells(1, ciItem).value
        If ColIndex(db, "Quantity") > 0 And ciQty > 0 Then newRow.Range.Cells(1, ColIndex(db, "Quantity")).value = srcR.Cells(1, ciQty).value
        If ColIndex(db, "UnitCost") > 0 And ciUnit > 0 Then newRow.Range.Cells(1, ColIndex(db, "UnitCost")).value = srcR.Cells(1, ciUnit).value

        ' Optionally compute TotalCost if present
        If ColIndex(db, "TotalCost") > 0 Then
            Dim q As Double, u As Double
            If ciQty > 0 Then q = Nz(srcR.Cells(1, ciQty).value, 0) Else q = 0
            If ciUnit > 0 Then u = Nz(srcR.Cells(1, ciUnit).value, 0) Else u = 0
            newRow.Range.Cells(1, ColIndex(db, "TotalCost")).value = q * u
        End If

        If ColIndex(db, "Supplier") > 0 And ciSupplier > 0 Then newRow.Range.Cells(1, ColIndex(db, "Supplier")).value = srcR.Cells(1, ciSupplier).value

        ' Audit
        AuditWrite "Create", "tblConsumables", IIf(dbCID > 0, newRow.Range.Cells(1, dbCID).value, "#?"), userName, "Imported from staging"

        ' Delete the staging row
        stg.ListRows(i).Delete

        CommitStgConsumables = CommitStgConsumables + 1
    Next i

ExitPoint:
    Exit Function
ErrHandler:
    MsgBox "Error in CommitStgConsumables: " & Err.Description, vbExclamation
    Resume ExitPoint
End Function

' ---------- Payments ----------
Public Function CommitStgPayments(ByVal projectID As Long, ByVal userName As String) As Long
    Dim stg As ListObject, db As ListObject
    Dim i As Long, newRow As ListRow, srcR As Range
    Set stg = GetTable("tblStgPayments")
    Set db = GetTable("tblPayments")
    If stg Is Nothing Or db Is Nothing Then Exit Function

    On Error GoTo ErrHandler
    For i = stg.ListRows.Count To 1 Step -1
        Set srcR = stg.DataBodyRange.rows(i)
        Set newRow = db.ListRows.Add
        ' safe mapping
        If ColIndex(db, "PaymentID") > 0 Then newRow.Range.Cells(1, ColIndex(db, "PaymentID")).value = NextID("tblPayments", "PaymentID")
        If ColIndex(db, "ProjectID") > 0 Then newRow.Range.Cells(1, ColIndex(db, "ProjectID")).value = projectID
        If ColIndex(db, "WorkerID") > 0 Then newRow.Range.Cells(1, ColIndex(db, "WorkerID")).value = Nz(srcR.Cells(1, ColIndex(stg, "WorkerID")).value, "")
        If ColIndex(db, "DatePaid") > 0 Then newRow.Range.Cells(1, ColIndex(db, "DatePaid")).value = Nz(srcR.Cells(1, ColIndex(stg, "DatePaid")).value, "")
        If ColIndex(db, "Hours") > 0 Then newRow.Range.Cells(1, ColIndex(db, "Hours")).value = Nz(srcR.Cells(1, ColIndex(stg, "Hours")).value, 0)
        If ColIndex(db, "Rate") > 0 Then newRow.Range.Cells(1, ColIndex(db, "Rate")).value = Nz(srcR.Cells(1, ColIndex(stg, "Rate")).value, 0)
        If ColIndex(db, "Amount") > 0 Then newRow.Range.Cells(1, ColIndex(db, "Amount")).value = Nz(srcR.Cells(1, ColIndex(stg, "Amount")).value, 0)
        If ColIndex(db, "PaymentMethodID") > 0 Then newRow.Range.Cells(1, ColIndex(db, "PaymentMethodID")).value = Nz(srcR.Cells(1, ColIndex(stg, "PaymentMethodID")).value, "")
        If ColIndex(db, "Notes") > 0 Then newRow.Range.Cells(1, ColIndex(db, "Notes")).value = Nz(srcR.Cells(1, ColIndex(stg, "Notes")).value, "")

        AuditWrite "Create", "tblPayments", IIf(ColIndex(db, "PaymentID") > 0, newRow.Range.Cells(1, ColIndex(db, "PaymentID")).value, "#?"), userName, "Imported from staging"

        stg.ListRows(i).Delete
        CommitStgPayments = CommitStgPayments + 1
    Next i

ExitPoint:
    Exit Function
ErrHandler:
    MsgBox "Error in CommitStgPayments: " & Err.Description, vbExclamation
    Resume ExitPoint
End Function

' ---------- Logistics ----------
Public Function CommitStgLogistics(ByVal projectID As Long, ByVal userName As String) As Long
    Dim stg As ListObject, db As ListObject
    Dim i As Long, newRow As ListRow, srcR As Range
    Set stg = GetTable("tblStgLogistics")
    Set db = GetTable("tblLogistics")
    If stg Is Nothing Or db Is Nothing Then Exit Function

    On Error GoTo ErrHandler
    For i = stg.ListRows.Count To 1 Step -1
        Set srcR = stg.DataBodyRange.rows(i)
        Set newRow = db.ListRows.Add
        If ColIndex(db, "LogisticID") > 0 Then newRow.Range.Cells(1, ColIndex(db, "LogisticID")).value = NextID("tblLogistics", "LogisticID")
        If ColIndex(db, "ProjectID") > 0 Then newRow.Range.Cells(1, ColIndex(db, "ProjectID")).value = projectID
        If ColIndex(db, "Date") > 0 Then newRow.Range.Cells(1, ColIndex(db, "Date")).value = Nz(srcR.Cells(1, ColIndex(stg, "Date")).value, "")
        If ColIndex(db, "CategoryID") > 0 Then newRow.Range.Cells(1, ColIndex(db, "CategoryID")).value = Nz(srcR.Cells(1, ColIndex(stg, "CategoryID")).value, "")
        If ColIndex(db, "Description") > 0 Then newRow.Range.Cells(1, ColIndex(db, "Description")).value = Nz(srcR.Cells(1, ColIndex(stg, "Description")).value, "")
        If ColIndex(db, "Amount") > 0 Then newRow.Range.Cells(1, ColIndex(db, "Amount")).value = Nz(srcR.Cells(1, ColIndex(stg, "Amount")).value, 0)
        If ColIndex(db, "Vendor") > 0 Then newRow.Range.Cells(1, ColIndex(db, "Vendor")).value = Nz(srcR.Cells(1, ColIndex(stg, "Vendor")).value, "")

        AuditWrite "Create", "tblLogistics", IIf(ColIndex(db, "LogisticID") > 0, newRow.Range.Cells(1, ColIndex(db, "LogisticID")).value, "#?"), userName, "Imported from staging"

        stg.ListRows(i).Delete
        CommitStgLogistics = CommitStgLogistics + 1
    Next i

ExitPoint:
    Exit Function
ErrHandler:
    MsgBox "Error in CommitStgLogistics: " & Err.Description, vbExclamation
    Resume ExitPoint
End Function

' ----------- Safety-----------------
Public Function CommitStgSafety(ByVal projectID As Long, ByVal userName As String) As Long
    Dim stg As ListObject, db As ListObject
    Dim i As Long, newRow As ListRow, srcR As Range
    Dim ciDate As Long, ciCategory As Long, ciItem As Long, ciQty As Long, ciUnit As Long, ciSupplier As Long, ciNotes As Long
    Dim dbSID As Long, dbPID As Long, dbDate As Long

    Set stg = GetTable("tblStgSafety")
    Set db = GetTable("tblSafety")
    If stg Is Nothing Or db Is Nothing Then Exit Function

    ' Cache column indexes (0 means missing)
    ciDate = ColIndex(stg, "Date")
    ciCategory = ColIndex(stg, "CategoryID")
    ciItem = ColIndex(stg, "ItemDescription")
    ciQty = ColIndex(stg, "Quantity")
    ciUnit = ColIndex(stg, "UnitCost")
    ciSupplier = ColIndex(stg, "Supplier")
    ciNotes = ColIndex(stg, "Notes")

    dbSID = ColIndex(db, "SafetyID")
    dbPID = ColIndex(db, "ProjectID")
    dbDate = ColIndex(db, "Date")

    On Error GoTo ErrHandler
    For i = stg.ListRows.Count To 1 Step -1
        Set srcR = stg.DataBodyRange.rows(i)

        ' Add row to DB
        Set newRow = db.ListRows.Add
        If dbSID > 0 Then newRow.Range.Cells(1, dbSID).value = NextID("tblSafety", "SafetyID")
        If dbPID > 0 Then newRow.Range.Cells(1, dbPID).value = projectID
        If dbDate > 0 And ciDate > 0 Then newRow.Range.Cells(1, dbDate).value = srcR.Cells(1, ciDate).value

        ' Safe writes for other fields if columns exist
        If ColIndex(db, "CategoryID") > 0 And ciCategory > 0 Then newRow.Range.Cells(1, ColIndex(db, "CategoryID")).value = srcR.Cells(1, ciCategory).value
        If ColIndex(db, "ItemDescription") > 0 And ciItem > 0 Then newRow.Range.Cells(1, ColIndex(db, "ItemDescription")).value = srcR.Cells(1, ciItem).value
        If ColIndex(db, "Quantity") > 0 And ciQty > 0 Then newRow.Range.Cells(1, ColIndex(db, "Quantity")).value = srcR.Cells(1, ciQty).value
        If ColIndex(db, "UnitCost") > 0 And ciUnit > 0 Then newRow.Range.Cells(1, ColIndex(db, "UnitCost")).value = srcR.Cells(1, ciUnit).value

        ' Optionally compute TotalCost if present
        If ColIndex(db, "TotalCost") > 0 Then
            Dim qS As Double, uS As Double
            If ciQty > 0 Then qS = Nz(srcR.Cells(1, ciQty).value, 0) Else qS = 0
            If ciUnit > 0 Then uS = Nz(srcR.Cells(1, ciUnit).value, 0) Else uS = 0
            newRow.Range.Cells(1, ColIndex(db, "TotalCost")).value = qS * uS
        End If

        If ColIndex(db, "Supplier") > 0 And ciSupplier > 0 Then newRow.Range.Cells(1, ColIndex(db, "Supplier")).value = srcR.Cells(1, ciSupplier).value
        If ColIndex(db, "Notes") > 0 And ciNotes > 0 Then newRow.Range.Cells(1, ColIndex(db, "Notes")).value = srcR.Cells(1, ciNotes).value

        ' Optional created metadata if columns exist
        If ColIndex(db, "CreatedBy") > 0 Then newRow.Range.Cells(1, ColIndex(db, "CreatedBy")).value = userName
        If ColIndex(db, "CreatedAt") > 0 Then newRow.Range.Cells(1, ColIndex(db, "CreatedAt")).value = Now

        ' Audit
        AuditWrite "Create", "tblSafety", IIf(dbSID > 0, newRow.Range.Cells(1, dbSID).value, "#?"), userName, "Imported from staging"

        ' Delete the staging row (delete by ListRows index to avoid messing enumeration)
        stg.ListRows(i).Delete

        CommitStgSafety = CommitStgSafety + 1
    Next i

ExitPoint:
    Exit Function
ErrHandler:
    MsgBox "Error in CommitStgSafety: " & Err.Description, vbExclamation
    Resume ExitPoint
End Function


'------------- Material ------------------
Public Function CommitStgMaterials(ByVal projectID As Long, ByVal userName As String) As Long
    Dim stg As ListObject, db As ListObject
    Dim i As Long, newRow As ListRow, srcR As Range
    Dim ciDate As Long, ciCategory As Long, ciItem As Long, ciQty As Long, ciUnit As Long, ciUnitCost As Long, ciSupplier As Long, ciNotes As Long
    Dim dbMID As Long, dbPID As Long, dbDate As Long

    Set stg = GetTable("tblStgMaterials")
    Set db = GetTable("tblMaterials")
    If stg Is Nothing Or db Is Nothing Then Exit Function

    ' Cache column indexes (0 means missing)
    ciDate = ColIndex(stg, "Date")
    ciCategory = ColIndex(stg, "CategoryID")
    ciItem = ColIndex(stg, "ItemDescription")
    ciQty = ColIndex(stg, "Quantity")
    ciUnit = ColIndex(stg, "Unit")
    ciUnitCost = ColIndex(stg, "UnitCost")
    ciSupplier = ColIndex(stg, "Supplier")
    ciNotes = ColIndex(stg, "Notes")

    dbMID = ColIndex(db, "MaterialID")
    dbPID = ColIndex(db, "ProjectID")
    dbDate = ColIndex(db, "Date")

    On Error GoTo ErrHandler
    For i = stg.ListRows.Count To 1 Step -1
        Set srcR = stg.DataBodyRange.rows(i)

        ' Add row to DB
        Set newRow = db.ListRows.Add
        If dbMID > 0 Then newRow.Range.Cells(1, dbMID).value = NextID("tblMaterials", "MaterialID")
        If dbPID > 0 Then newRow.Range.Cells(1, dbPID).value = projectID
        If dbDate > 0 And ciDate > 0 Then newRow.Range.Cells(1, dbDate).value = srcR.Cells(1, ciDate).value

        ' Safe writes for other fields if columns exist
        If ColIndex(db, "CategoryID") > 0 And ciCategory > 0 Then newRow.Range.Cells(1, ColIndex(db, "CategoryID")).value = srcR.Cells(1, ciCategory).value
        If ColIndex(db, "ItemDescription") > 0 And ciItem > 0 Then newRow.Range.Cells(1, ColIndex(db, "ItemDescription")).value = srcR.Cells(1, ciItem).value
        If ColIndex(db, "Quantity") > 0 And ciQty > 0 Then newRow.Range.Cells(1, ColIndex(db, "Quantity")).value = srcR.Cells(1, ciQty).value
        If ColIndex(db, "Unit") > 0 And ciUnit > 0 Then newRow.Range.Cells(1, ColIndex(db, "Unit")).value = srcR.Cells(1, ciUnit).value
        If ColIndex(db, "UnitCost") > 0 And ciUnitCost > 0 Then newRow.Range.Cells(1, ColIndex(db, "UnitCost")).value = srcR.Cells(1, ciUnitCost).value

        ' Optionally compute TotalCost if present
        If ColIndex(db, "TotalCost") > 0 Then
            Dim qM As Double, uM As Double
            If ciQty > 0 Then qM = Nz(srcR.Cells(1, ciQty).value, 0) Else qM = 0
            If ciUnitCost > 0 Then uM = Nz(srcR.Cells(1, ciUnitCost).value, 0) Else uM = 0
            newRow.Range.Cells(1, ColIndex(db, "TotalCost")).value = qM * uM
        End If

        If ColIndex(db, "Supplier") > 0 And ciSupplier > 0 Then newRow.Range.Cells(1, ColIndex(db, "Supplier")).value = srcR.Cells(1, ciSupplier).value
        If ColIndex(db, "Notes") > 0 And ciNotes > 0 Then newRow.Range.Cells(1, ColIndex(db, "Notes")).value = srcR.Cells(1, ciNotes).value

        ' Optional created metadata if columns exist
        If ColIndex(db, "CreatedBy") > 0 Then newRow.Range.Cells(1, ColIndex(db, "CreatedBy")).value = userName
        If ColIndex(db, "CreatedAt") > 0 Then newRow.Range.Cells(1, ColIndex(db, "CreatedAt")).value = Now

        ' Audit
        AuditWrite "Create", "tblMaterials", IIf(dbMID > 0, newRow.Range.Cells(1, dbMID).value, "#?"), userName, "Imported from staging"

        ' Delete staging row
        stg.ListRows(i).Delete

        CommitStgMaterials = CommitStgMaterials + 1
    Next i

ExitPoint:
    Exit Function
ErrHandler:
    MsgBox "Error in CommitStgMaterials: " & Err.Description, vbExclamation
    Resume ExitPoint
End Function



' Nz helper
Private Function Nz(v, Optional defaultVal = 0)
    If IsError(v) Then Nz = defaultVal: Exit Function
    If IsEmpty(v) Then Nz = defaultVal: Exit Function
    If Len(Trim(CStr(v))) = 0 Then Nz = defaultVal Else Nz = v
End Function


' Return 0 if not found, otherwise the 1-based column index
Public Function ColIndex(lo As ListObject, colName As String) As Long
    On Error Resume Next
    If lo Is Nothing Then
        ColIndex = 0
    Else
        ColIndex = lo.ListColumns(colName).Index
        If Err.Number <> 0 Then
            ColIndex = 0
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

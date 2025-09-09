Attribute VB_Name = "modDBActions"
Option Explicit

' ---------- Consumables ----------
Public Sub AddConsumableToDB(projectID As Long, dte As Variant, category As String, desc As String, qty As Double, unitCost As Double, supplier As String, userName As String)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblConsumables")
    If lo Is Nothing Then MsgBox "tblConsumables missing": Exit Sub
    Set lr = lo.ListRows.Add
    lr.Range(lo.ListColumns("ConsumableID").Index).value = NextID("tblConsumables", "ConsumableID")
    lr.Range(lo.ListColumns("ProjectID").Index).value = projectID
    lr.Range(lo.ListColumns("Date").Index).value = CDate(dte)
    lr.Range(lo.ListColumns("CategoryID").Index).value = category
    lr.Range(lo.ListColumns("ItemDescription").Index).value = desc
    lr.Range(lo.ListColumns("Quantity").Index).value = qty
    lr.Range(lo.ListColumns("UnitCost").Index).value = unitCost
    lr.Range(lo.ListColumns("TotalCost").Index).value = qty * unitCost
    lr.Range(lo.ListColumns("Supplier").Index).value = supplier
    AuditWrite "Create", "tblConsumables", lr.Range(lo.ListColumns("ConsumableID").Index).value, userName, "Added to project " & projectID
End Sub

Public Sub UpdateConsumable(consumableID As Long, dte As Variant, category As String, desc As String, qty As Double, unitCost As Double, supplier As String, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblConsumables")
    If lo Is Nothing Then Exit Sub
    For Each r In lo.ListColumns("ConsumableID").DataBodyRange.rows
        If r.value = consumableID Then
            r.EntireRow.Cells(1, lo.ListColumns("Date").Index).value = CDate(dte)
            r.EntireRow.Cells(1, lo.ListColumns("CategoryID").Index).value = category
            r.EntireRow.Cells(1, lo.ListColumns("ItemDescription").Index).value = desc
            r.EntireRow.Cells(1, lo.ListColumns("Quantity").Index).value = qty
            r.EntireRow.Cells(1, lo.ListColumns("UnitCost").Index).value = unitCost
            r.EntireRow.Cells(1, lo.ListColumns("TotalCost").Index).value = qty * unitCost
            r.EntireRow.Cells(1, lo.ListColumns("Supplier").Index).value = supplier
            AuditWrite "Update", "tblConsumables", consumableID, userName, "Updated"
            Exit For
        End If
    Next r
End Sub

Public Sub DeleteConsumable(consumableID As Long, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblConsumables")
    If lo Is Nothing Then Exit Sub
    For Each r In lo.ListColumns("ConsumableID").DataBodyRange.rows
        If r.value = consumableID Then
            r.EntireRow.Delete
            AuditWrite "Delete", "tblConsumables", consumableID, userName, "Deleted"
            Exit For
        End If
    Next r
End Sub

' ---------- Payments ----------
Public Sub AddPaymentToDB(projectID As Long, workerID As Variant, datePaid As Variant, hours As Double, rate As Double, amount As Double, paymentMethod As String, notes As String, userName As String)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblPayments")
    If lo Is Nothing Then MsgBox "tblPayments missing": Exit Sub
    Set lr = lo.ListRows.Add
    lr.Range(lo.ListColumns("PaymentID").Index).value = NextID("tblPayments", "PaymentID")
    lr.Range(lo.ListColumns("ProjectID").Index).value = projectID
    lr.Range(lo.ListColumns("WorkerID").Index).value = workerID
    lr.Range(lo.ListColumns("DatePaid").Index).value = CDate(datePaid)
    lr.Range(lo.ListColumns("Hours").Index).value = hours
    lr.Range(lo.ListColumns("Rate").Index).value = rate
    lr.Range(lo.ListColumns("Amount").Index).value = amount
    lr.Range(lo.ListColumns("PaymentMethodID").Index).value = paymentMethod
    lr.Range(lo.ListColumns("Notes").Index).value = notes
    AuditWrite "Create", "tblPayments", lr.Range(lo.ListColumns("PaymentID").Index).value, userName, "Added to project " & projectID
End Sub

Public Sub UpdatePayment(paymentID As Long, workerID As Variant, datePaid As Variant, hours As Double, rate As Double, amount As Double, paymentMethod As String, notes As String, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblPayments")
    If lo Is Nothing Then Exit Sub
    For Each r In lo.ListColumns("PaymentID").DataBodyRange.rows
        If r.value = paymentID Then
            r.EntireRow.Cells(1, lo.ListColumns("WorkerID").Index).value = workerID
            r.EntireRow.Cells(1, lo.ListColumns("DatePaid").Index).value = CDate(datePaid)
            r.EntireRow.Cells(1, lo.ListColumns("Hours").Index).value = hours
            r.EntireRow.Cells(1, lo.ListColumns("Rate").Index).value = rate
            r.EntireRow.Cells(1, lo.ListColumns("Amount").Index).value = amount
            r.EntireRow.Cells(1, lo.ListColumns("PaymentMethodID").Index).value = paymentMethod
            r.EntireRow.Cells(1, lo.ListColumns("Notes").Index).value = notes
            AuditWrite "Update", "tblPayments", paymentID, userName, "Updated"
            Exit For
        End If
    Next r
End Sub

Public Sub DeletePayment(paymentID As Long, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblPayments")
    If lo Is Nothing Then Exit Sub
    For Each r In lo.ListColumns("PaymentID").DataBodyRange.rows
        If r.value = paymentID Then
            r.EntireRow.Delete
            AuditWrite "Delete", "tblPayments", paymentID, userName, "Deleted"
            Exit For
        End If
    Next r
End Sub

' ---------- Logistics ----------
Public Sub AddLogisticToDB(projectID As Long, dte As Variant, category As String, desc As String, amount As Double, vendor As String, userName As String)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblLogistics")
    If lo Is Nothing Then MsgBox "tblLogistics missing": Exit Sub
    Set lr = lo.ListRows.Add
    lr.Range(lo.ListColumns("LogisticID").Index).value = NextID("tblLogistics", "LogisticID")
    lr.Range(lo.ListColumns("ProjectID").Index).value = projectID
    lr.Range(lo.ListColumns("Date").Index).value = CDate(dte)
    lr.Range(lo.ListColumns("CategoryID").Index).value = category
    lr.Range(lo.ListColumns("Description").Index).value = desc
    lr.Range(lo.ListColumns("Amount").Index).value = amount
    lr.Range(lo.ListColumns("Vendor").Index).value = vendor
    AuditWrite "Create", "tblLogistics", lr.Range(lo.ListColumns("LogisticID").Index).value, userName, "Added to project " & projectID
End Sub

Public Sub UpdateLogistic(logID As Long, dte As Variant, category As String, desc As String, amount As Double, vendor As String, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblLogistics")
    If lo Is Nothing Then Exit Sub
    For Each r In lo.ListColumns("LogisticID").DataBodyRange.rows
        If r.value = logID Then
            r.EntireRow.Cells(1, lo.ListColumns("Date").Index).value = CDate(dte)
            r.EntireRow.Cells(1, lo.ListColumns("CategoryID").Index).value = category
            r.EntireRow.Cells(1, lo.ListColumns("Description").Index).value = desc
            r.EntireRow.Cells(1, lo.ListColumns("Amount").Index).value = amount
            r.EntireRow.Cells(1, lo.ListColumns("Vendor").Index).value = vendor
            AuditWrite "Update", "tblLogistics", logID, userName, "Updated"
            Exit For
        End If
    Next r
End Sub

Public Sub DeleteLogistic(logID As Long, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblLogistics")
    If lo Is Nothing Then Exit Sub
    For Each r In lo.ListColumns("LogisticID").DataBodyRange.rows
        If r.value = logID Then
            r.EntireRow.Delete
            AuditWrite "Delete", "tblLogistics", logID, userName, "Deleted"
            Exit For
        End If
    Next r
End Sub


' ---------- Safety ----------
Public Sub AddSafetyToDB(projectID As Long, dte As Variant, category As String, desc As String, qty As Double, unitCost As Double, supplier As String, notes As String, userName As String)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblSafety")
    If lo Is Nothing Then MsgBox "tblSafety missing": Exit Sub

    Set lr = lo.ListRows.Add
    lr.Range(lo.ListColumns("SafetyID").Index).value = NextID("tblSafety", "SafetyID")
    lr.Range(lo.ListColumns("ProjectID").Index).value = projectID
    lr.Range(lo.ListColumns("Date").Index).value = CDate(dte)
    lr.Range(lo.ListColumns("CategoryID").Index).value = category
    lr.Range(lo.ListColumns("ItemDescription").Index).value = desc
    lr.Range(lo.ListColumns("Quantity").Index).value = qty
    lr.Range(lo.ListColumns("UnitCost").Index).value = unitCost

    ' Set TotalCost if column exists
    If ColIndex(lo, "TotalCost") > 0 Then
        lr.Range(lo.ListColumns("TotalCost").Index).value = qty * unitCost
    End If

    lr.Range(lo.ListColumns("Supplier").Index).value = supplier
    If ColIndex(lo, "Notes") > 0 Then lr.Range(lo.ListColumns("Notes").Index).value = notes

    ' Optional created metadata if columns present
    If ColIndex(lo, "CreatedBy") > 0 Then lr.Range(lo.ListColumns("CreatedBy").Index).value = userName
    If ColIndex(lo, "CreatedAt") > 0 Then lr.Range(lo.ListColumns("CreatedAt").Index).value = Now

    AuditWrite "Create", "tblSafety", lr.Range(lo.ListColumns("SafetyID").Index).value, userName, "Added to project " & projectID
End Sub

Public Sub UpdateSafety(safetyID As Long, dte As Variant, category As String, desc As String, qty As Double, unitCost As Double, supplier As String, notes As String, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblSafety")
    If lo Is Nothing Then Exit Sub
    If lo.ListColumns("SafetyID").DataBodyRange Is Nothing Then Exit Sub

    For Each r In lo.ListColumns("SafetyID").DataBodyRange.rows
        If r.value = safetyID Then
            r.EntireRow.Cells(1, lo.ListColumns("Date").Index).value = CDate(dte)
            r.EntireRow.Cells(1, lo.ListColumns("CategoryID").Index).value = category
            r.EntireRow.Cells(1, lo.ListColumns("ItemDescription").Index).value = desc
            r.EntireRow.Cells(1, lo.ListColumns("Quantity").Index).value = qty
            r.EntireRow.Cells(1, lo.ListColumns("UnitCost").Index).value = unitCost

            ' Update TotalCost if column exists
            If ColIndex(lo, "TotalCost") > 0 Then
                r.EntireRow.Cells(1, lo.ListColumns("TotalCost").Index).value = qty * unitCost
            End If

            r.EntireRow.Cells(1, lo.ListColumns("Supplier").Index).value = supplier
            If ColIndex(lo, "Notes") > 0 Then r.EntireRow.Cells(1, lo.ListColumns("Notes").Index).value = notes

            If ColIndex(lo, "CreatedBy") > 0 Then r.EntireRow.Cells(1, lo.ListColumns("CreatedBy").Index).value = userName
            If ColIndex(lo, "CreatedAt") > 0 Then r.EntireRow.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now

            AuditWrite "Update", "tblSafety", safetyID, userName, "Updated"
            Exit For
        End If
    Next r
End Sub

Public Sub DeleteSafety(safetyID As Long, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblSafety")
    If lo Is Nothing Then Exit Sub
    If lo.ListColumns("SafetyID").DataBodyRange Is Nothing Then Exit Sub

    For Each r In lo.ListColumns("SafetyID").DataBodyRange.rows
        If r.value = safetyID Then
            r.EntireRow.Delete
            AuditWrite "Delete", "tblSafety", safetyID, userName, "Deleted"
            Exit For
        End If
    Next r
End Sub

' ---------- Materials ----------
Public Sub AddMaterialToDB(projectID As Long, dte As Variant, category As String, desc As String, qty As Double, unit As String, unitCost As Double, supplier As String, notes As String, userName As String)
    Dim lo As ListObject, lr As ListRow
    Set lo = GetTable("tblMaterials")
    If lo Is Nothing Then MsgBox "tblMaterials missing": Exit Sub

    Set lr = lo.ListRows.Add
    lr.Range(lo.ListColumns("MaterialID").Index).value = NextID("tblMaterials", "MaterialID")
    lr.Range(lo.ListColumns("ProjectID").Index).value = projectID
    lr.Range(lo.ListColumns("Date").Index).value = CDate(dte)
    lr.Range(lo.ListColumns("CategoryID").Index).value = category
    lr.Range(lo.ListColumns("ItemDescription").Index).value = desc
    lr.Range(lo.ListColumns("Quantity").Index).value = qty
    lr.Range(lo.ListColumns("Unit").Index).value = unit
    lr.Range(lo.ListColumns("UnitCost").Index).value = unitCost

    ' Set TotalCost if column exists
    If ColIndex(lo, "TotalCost") > 0 Then
        lr.Range(lo.ListColumns("TotalCost").Index).value = qty * unitCost
    End If

    lr.Range(lo.ListColumns("Supplier").Index).value = supplier
    If ColIndex(lo, "Notes") > 0 Then lr.Range(lo.ListColumns("Notes").Index).value = notes

    ' Optional created metadata if columns present
    If ColIndex(lo, "CreatedBy") > 0 Then lr.Range(lo.ListColumns("CreatedBy").Index).value = userName
    If ColIndex(lo, "CreatedAt") > 0 Then lr.Range(lo.ListColumns("CreatedAt").Index).value = Now

    AuditWrite "Create", "tblMaterials", lr.Range(lo.ListColumns("MaterialID").Index).value, userName, "Added to project " & projectID
End Sub

Public Sub UpdateMaterial(materialID As Long, dte As Variant, category As String, desc As String, qty As Double, unit As String, unitCost As Double, supplier As String, notes As String, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblMaterials")
    If lo Is Nothing Then Exit Sub
    If lo.ListColumns("MaterialID").DataBodyRange Is Nothing Then Exit Sub

    For Each r In lo.ListColumns("MaterialID").DataBodyRange.rows
        If r.value = materialID Then
            r.EntireRow.Cells(1, lo.ListColumns("Date").Index).value = CDate(dte)
            r.EntireRow.Cells(1, lo.ListColumns("CategoryID").Index).value = category
            r.EntireRow.Cells(1, lo.ListColumns("ItemDescription").Index).value = desc
            r.EntireRow.Cells(1, lo.ListColumns("Quantity").Index).value = qty
            r.EntireRow.Cells(1, lo.ListColumns("Unit").Index).value = unit
            r.EntireRow.Cells(1, lo.ListColumns("UnitCost").Index).value = unitCost

            ' Update TotalCost if column exists
            If ColIndex(lo, "TotalCost") > 0 Then
                r.EntireRow.Cells(1, lo.ListColumns("TotalCost").Index).value = qty * unitCost
            End If

            r.EntireRow.Cells(1, lo.ListColumns("Supplier").Index).value = supplier
            If ColIndex(lo, "Notes") > 0 Then r.EntireRow.Cells(1, lo.ListColumns("Notes").Index).value = notes

            If ColIndex(lo, "CreatedBy") > 0 Then r.EntireRow.Cells(1, lo.ListColumns("CreatedBy").Index).value = userName
            If ColIndex(lo, "CreatedAt") > 0 Then r.EntireRow.Cells(1, lo.ListColumns("CreatedAt").Index).value = Now

            AuditWrite "Update", "tblMaterials", materialID, userName, "Updated"
            Exit For
        End If
    Next r
End Sub

Public Sub DeleteMaterial(materialID As Long, userName As String)
    Dim lo As ListObject, r As Range
    Set lo = GetTable("tblMaterials")
    If lo Is Nothing Then Exit Sub
    If lo.ListColumns("MaterialID").DataBodyRange Is Nothing Then Exit Sub

    For Each r In lo.ListColumns("MaterialID").DataBodyRange.rows
        If r.value = materialID Then
            r.EntireRow.Delete
            AuditWrite "Delete", "tblMaterials", materialID, userName, "Deleted"
            Exit For
        End If
    Next r
End Sub


' ---------- Populate listboxes (DB) ----------
' Example: PopulateConsumablesListbox (DB)
Public Sub PopulateConsumablesListbox(frm As Object, projectID As Long)
    Dim lo As ListObject, i As Long
    Set lo = GetTable("tblConsumables")
    frm.lstConsumables.Clear
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ProjectID").Index).value = projectID Then
            frm.lstConsumables.AddItem
            frm.lstConsumables.List(frm.lstConsumables.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ConsumableID").Index).value
            frm.lstConsumables.List(frm.lstConsumables.ListCount - 1, 1) = Format(lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Date").Index).value, "yyyy-mm-dd")
            frm.lstConsumables.List(frm.lstConsumables.ListCount - 1, 2) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ItemDescription").Index).value
            frm.lstConsumables.List(frm.lstConsumables.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Quantity").Index).value
            frm.lstConsumables.List(frm.lstConsumables.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("UnitCost").Index).value
            frm.lstConsumables.List(frm.lstConsumables.ListCount - 1, 5) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Supplier").Index).value
        End If
    Next i
    ' hide first column
    frm.lstConsumables.ColumnCount = 6
    frm.lstConsumables.ColumnWidths = "0;80;220;50;70;120"
End Sub


Public Sub PopulatePaymentsListbox(frm As Object, projectID As Long)
    Dim lo As ListObject, i As Long
    Set lo = GetTable("tblPayments")
    frm.lstPayments.Clear
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ProjectID").Index).value = projectID Then
            frm.lstPayments.AddItem
            frm.lstPayments.List(frm.lstPayments.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("PaymentID").Index).value
            frm.lstPayments.List(frm.lstPayments.ListCount - 1, 1) = Format(lo.DataBodyRange.rows(i).Cells(lo.ListColumns("DatePaid").Index).value, "yyyy-mm-dd")
            frm.lstPayments.List(frm.lstPayments.ListCount - 1, 2) = GetWorkerNameByID(lo.DataBodyRange.rows(i).Cells(lo.ListColumns("WorkerID").Index).value)
            frm.lstPayments.List(frm.lstPayments.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Hours").Index).value
            frm.lstPayments.List(frm.lstPayments.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Amount").Index).value
            frm.lstPayments.List(frm.lstPayments.ListCount - 1, 5) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("PaymentMethodID").Index).value
        End If
    Next i
    ' hide first column
    frm.lstPayments.ColumnCount = 6
    frm.lstPayments.ColumnWidths = "0;80;200;60;80;120"
End Sub

Public Sub PopulateLogisticsListbox(frm As Object, projectID As Long)
    Dim lo As ListObject, i As Long
    Set lo = GetTable("tblLogistics")
    frm.lstLogistics.Clear
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ProjectID").Index).value = projectID Then
            frm.lstLogistics.AddItem
            frm.lstLogistics.List(frm.lstLogistics.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("LogisticID").Index).value
            frm.lstLogistics.List(frm.lstLogistics.ListCount - 1, 1) = Format(lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Date").Index).value, "yyyy-mm-dd")
            frm.lstLogistics.List(frm.lstLogistics.ListCount - 1, 2) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Description").Index).value
            frm.lstLogistics.List(frm.lstLogistics.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Amount").Index).value
            frm.lstLogistics.List(frm.lstLogistics.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Vendor").Index).value
        End If
    Next i
    ' hide first column
    frm.lstLogistics.ColumnCount = 5
    frm.lstLogistics.ColumnWidths = "0;80;220;120;120"
End Sub


Public Sub PopulateSafetyListbox(frm As Object, projectID As Long)
    Dim lo As ListObject, i As Long
    Set lo = GetTable("tblSafety")
    frm.lstSafety.Clear
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ProjectID").Index).value = projectID Then
            frm.lstSafety.AddItem
            frm.lstSafety.List(frm.lstSafety.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("SafetyID").Index).value
            frm.lstSafety.List(frm.lstSafety.ListCount - 1, 1) = Format(lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Date").Index).value, "yyyy-mm-dd")
            frm.lstSafety.List(frm.lstSafety.ListCount - 1, 2) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ItemDescription").Index).value
            frm.lstSafety.List(frm.lstSafety.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Quantity").Index).value
            frm.lstSafety.List(frm.lstSafety.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Unitcost").Index).value
            frm.lstSafety.List(frm.lstSafety.ListCount - 1, 5) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Supplier").Index).value
        End If
    Next i
    ' hide first column
    frm.lstSafety.ColumnCount = 6
    frm.lstSafety.ColumnWidths = "0;80;200;60;80;120"
End Sub

Public Sub PopulateMaterialsListbox(frm As Object, projectID As Long)
    Dim lo As ListObject, i As Long
    Set lo = GetTable("tblMaterials")
    frm.lstMaterials.Clear
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ProjectID").Index).value = projectID Then
            frm.lstMaterials.AddItem
            frm.lstMaterials.List(frm.lstMaterials.ListCount - 1, 0) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("MaterialID").Index).value
            frm.lstMaterials.List(frm.lstMaterials.ListCount - 1, 1) = Format(lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Date").Index).value, "yyyy-mm-dd")
            frm.lstMaterials.List(frm.lstMaterials.ListCount - 1, 2) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("ItemDescription").Index).value
            frm.lstMaterials.List(frm.lstMaterials.ListCount - 1, 3) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Quantity").Index).value
            frm.lstMaterials.List(frm.lstMaterials.ListCount - 1, 4) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Unit").Index).value
            frm.lstMaterials.List(frm.lstMaterials.ListCount - 1, 5) = lo.DataBodyRange.rows(i).Cells(lo.ListColumns("Unitcost").Index).value
        End If
    Next i
    ' hide first column
    frm.lstMaterials.ColumnCount = 6
    frm.lstMaterials.ColumnWidths = "0;80;200;60;80;120"
End Sub


Attribute VB_Name = "modReports"
Option Explicit

' --- Public API: generate report for ProjectID with filters ---
Public Sub GenerateProjectReport(ByVal projectID As Long, _
                                         Optional includeCons As Boolean = True, _
                                         Optional includePays As Boolean = True, _
                                         Optional includeLogs As Boolean = True, _
                                         Optional includeSafe As Boolean = True, _
                                         Optional includeMat As Boolean = True, _
                                         Optional dtFrom As Variant, _
                                         Optional dtTo As Variant, _
                                         Optional categoryFilter As String = "")

    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Rpt_Project")
    Dim r As Long: r = 1 ' row pointer
    Dim lo As ListObject, cRow As Range
    Dim totalQty As Double, totalAmount As Double
    Dim totalHours As Double, totalPay As Double
    Dim totalLog As Double
    Dim totalMatQty As Double, totalMat As Double
    Dim totalSafeQty As Double, totalSafe As Double
    Dim wasVisible As XlSheetVisibility
    wasVisible = ws.Visible
    ws.Visible = xlSheetVisible
    ws.Cells.Clear

    ' === HEADER INFO ===
    Dim projLo As ListObject: Set projLo = GetTable("tblProjects")
    Dim projRow As Range
    Set projRow = FindRowByID(projLo, "ProjectID", projectID)
    If projRow Is Nothing Then
        ws.Cells(1, 1).value = "Project not found (ID: " & projectID & ")"
        GoTo Done
    End If

    ws.Cells(r, 1).value = "Project Report": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    ws.Cells(r, 1).value = "Project ID: " & projectID: r = r + 1
    ws.Cells(r, 1).value = "Project Code: " & projRow.Cells(1, ColIndex(projLo, "ProjectCode")).value: r = r + 1
    ws.Cells(r, 1).value = "Project Name: " & projRow.Cells(1, ColIndex(projLo, "ProjectName")).value: r = r + 1
    ws.Cells(r, 1).value = "Client: " & projRow.Cells(1, ColIndex(projLo, "CompanyID")).value: r = r + 1
    ws.Cells(r, 1).value = "Date Range: " & IIf(IsDate(dtFrom), dtFrom, "ALL") & " to " & IIf(IsDate(dtTo), dtTo, "ALL"): r = r + 2

    ' === CONSUMABLES ===
    If includeCons Then
        Set lo = GetTable("tblConsumables")
        If Not lo Is Nothing Then
            ws.Cells(r, 1).value = "Consumables": ws.Cells(r, 1).Font.Bold = True: r = r + 1
            Dim hdrCons As Variant: hdrCons = Array("Date", "Category", "Item", "Qty", "UnitCost", "Total")
            Dim c As Long
            For c = 0 To UBound(hdrCons)
                ws.Cells(r, c + 1).value = hdrCons(c)
                ws.Cells(r, c + 1).Font.Bold = True
                ws.Cells(r, c + 1).HorizontalAlignment = xlCenter
            Next c
            r = r + 1
            totalQty = 0: totalAmount = 0
            For Each cRow In lo.DataBodyRange.rows
                If cRow.Cells(1, ColIndex(lo, "ProjectID")).value = projectID Then
                    Dim entryDate As Variant: entryDate = cRow.Cells(1, ColIndex(lo, "Date")).value
                    If PassesDateFilter(entryDate, dtFrom, dtTo) Then
                        If categoryFilter = "" Or InStr(1, CStr(cRow.Cells(1, ColIndex(lo, "CategoryID")).value), categoryFilter, vbTextCompare) > 0 Then
                            ws.Cells(r, 1).value = entryDate
                            ws.Cells(r, 2).value = cRow.Cells(1, ColIndex(lo, "CategoryID")).value
                            ws.Cells(r, 3).value = cRow.Cells(1, ColIndex(lo, "ItemDescription")).value
                            ws.Cells(r, 4).value = cRow.Cells(1, ColIndex(lo, "Quantity")).value
                            ws.Cells(r, 5).value = cRow.Cells(1, ColIndex(lo, "UnitCost")).value
                            ws.Cells(r, 6).value = ws.Cells(r, 4).value * ws.Cells(r, 5).value
                            
                            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).HorizontalAlignment = xlLeft
                            ws.Cells(r, 4).HorizontalAlignment = xlLeft
                            ws.Range(ws.Cells(r, 5), ws.Cells(r, 6)).HorizontalAlignment = xlLeft
                            totalQty = totalQty + ws.Cells(r, 4).value
                            totalAmount = totalAmount + ws.Cells(r, 6).value
                            r = r + 1
                        End If
                    End If
                End If
            Next cRow
            ' Totals row
            ws.Cells(r, 3).value = "TOTAL": ws.Cells(r, 3).Font.Bold = True: ws.Cells(r, 3).HorizontalAlignment = xlRight
            ws.Cells(r, 4).value = totalQty: ws.Cells(r, 4).Font.Bold = True
            ws.Cells(r, 6).value = totalAmount: ws.Cells(r, 6).Font.Bold = True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).HorizontalAlignment = xlLeft
            r = r + 2
        End If
    End If

    ' === PAYMENTS ===
    If includePays Then
        Set lo = GetTable("tblPayments")
        If Not lo Is Nothing Then
            ws.Cells(r, 1).value = "Payments": ws.Cells(r, 1).Font.Bold = True: r = r + 1
            Dim hdrPay As Variant: hdrPay = Array("Date", "Worker", "Hours", "Rate", "Amount")
            For c = 0 To UBound(hdrPay)
                ws.Cells(r, c + 1).value = hdrPay(c)
                ws.Cells(r, c + 1).Font.Bold = True
                ws.Cells(r, c + 1).HorizontalAlignment = xlCenter
            Next c
            r = r + 1
            totalHours = 0: totalPay = 0
            For Each cRow In lo.DataBodyRange.rows
                If cRow.Cells(1, ColIndex(lo, "ProjectID")).value = projectID Then
                    Dim pDate As Variant: pDate = cRow.Cells(1, ColIndex(lo, "DatePaid")).value
                    If PassesDateFilter(pDate, dtFrom, dtTo) Then
                        ws.Cells(r, 1).value = pDate
                        ws.Cells(r, 2).value = GetWorkerNameByID(cRow.Cells(1, ColIndex(lo, "WorkerID")).value)
                        ws.Cells(r, 3).value = cRow.Cells(1, ColIndex(lo, "Hours")).value
                        ws.Cells(r, 4).value = cRow.Cells(1, ColIndex(lo, "Rate")).value
                        ws.Cells(r, 5).value = cRow.Cells(1, ColIndex(lo, "Amount")).value
                        
                        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).HorizontalAlignment = xlLeft
                        ws.Cells(r, 3).HorizontalAlignment = xlLeft
                        ws.Range(ws.Cells(r, 4), ws.Cells(r, 5)).HorizontalAlignment = xlLeft
                        
                        totalHours = totalHours + ws.Cells(r, 3).value
                        totalPay = totalPay + ws.Cells(r, 5).value
                        r = r + 1
                    End If
                End If
            Next cRow
            ' Totals row
            ws.Cells(r, 2).value = "TOTAL": ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).HorizontalAlignment = xlRight
            ws.Cells(r, 3).value = totalHours: ws.Cells(r, 3).Font.Bold = True
            ws.Cells(r, 5).value = totalPay: ws.Cells(r, 5).Font.Bold = True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).HorizontalAlignment = xlLeft
            r = r + 2
        End If
    End If

    ' === LOGISTICS ===
    If includeLogs Then
        Set lo = GetTable("tblLogistics")
        If Not lo Is Nothing Then
            ws.Cells(r, 1).value = "Logistics": ws.Cells(r, 1).Font.Bold = True: r = r + 1
            Dim hdrLog As Variant: hdrLog = Array("Date", "Category", "Description", "Vendor", "Amount")
            For c = 0 To UBound(hdrLog)
                ws.Cells(r, c + 1).value = hdrLog(c)
                ws.Cells(r, c + 1).Font.Bold = True
                ws.Cells(r, c + 1).HorizontalAlignment = xlCenter
            Next c
            r = r + 1
            totalLog = 0
            For Each cRow In lo.DataBodyRange.rows
                If cRow.Cells(1, ColIndex(lo, "ProjectID")).value = projectID Then
                    Dim lDate As Variant: lDate = cRow.Cells(1, ColIndex(lo, "Date")).value
                    If PassesDateFilter(lDate, dtFrom, dtTo) Then
                        ws.Cells(r, 1).value = lDate
                        ws.Cells(r, 2).value = cRow.Cells(1, ColIndex(lo, "CategoryID")).value
                        ws.Cells(r, 3).value = cRow.Cells(1, ColIndex(lo, "Description")).value
                        ws.Cells(r, 4).value = cRow.Cells(1, ColIndex(lo, "Vendor")).value
                        ws.Cells(r, 5).value = cRow.Cells(1, ColIndex(lo, "Amount")).value
                        
                        ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).HorizontalAlignment = xlLeft
                        ws.Cells(r, 5).HorizontalAlignment = xlLeft
                      
                        totalLog = totalLog + ws.Cells(r, 5).value
                        r = r + 1
                    End If
                End If
            Next cRow
            ' Totals row
            ws.Cells(r, 3).value = "TOTAL": ws.Cells(r, 3).Font.Bold = True: ws.Cells(r, 3).HorizontalAlignment = xlRight
            ws.Cells(r, 5).value = totalLog: ws.Cells(r, 5).Font.Bold = True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).HorizontalAlignment = xlLeft
            r = r + 2
        End If
    End If


    ' === SAFETY ===
    If includeSafe Then
        Set lo = GetTable("tblSafety")
        If Not lo Is Nothing Then
            ws.Cells(r, 1).value = "Safety Items": ws.Cells(r, 1).Font.Bold = True: r = r + 1
            Dim hdrSafe As Variant: hdrSafe = Array("Date", "CategoryID", "ItemDescription", "Quantity", "TotalCost")
            For c = 0 To UBound(hdrSafe)
                ws.Cells(r, c + 1).value = hdrSafe(c)
                ws.Cells(r, c + 1).Font.Bold = True
                ws.Cells(r, c + 1).HorizontalAlignment = xlCenter
            Next c
            r = r + 1
            totalLog = 0
            For Each cRow In lo.DataBodyRange.rows
                If cRow.Cells(1, ColIndex(lo, "ProjectID")).value = projectID Then
                    Dim sDate As Variant: sDate = cRow.Cells(1, ColIndex(lo, "Date")).value
                    If PassesDateFilter(sDate, dtFrom, dtTo) Then
                        ws.Cells(r, 1).value = sDate
                        ws.Cells(r, 2).value = cRow.Cells(1, ColIndex(lo, "CategoryID")).value
                        ws.Cells(r, 3).value = cRow.Cells(1, ColIndex(lo, "ItemDescription")).value
                        ws.Cells(r, 4).value = cRow.Cells(1, ColIndex(lo, "Quantity")).value
                        ws.Cells(r, 5).value = cRow.Cells(1, ColIndex(lo, "TotalCost")).value
                        
                        ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).HorizontalAlignment = xlLeft
                        ws.Cells(r, 4).HorizontalAlignment = xlLeft
                        ws.Cells(r, 5).HorizontalAlignment = xlLeft
                        
                        totalSafeQty = totalSafeQty + ws.Cells(r, 4).value
                        totalSafe = totalSafe + ws.Cells(r, 5).value
                        r = r + 1
                    End If
                End If
            Next cRow
            ' Totals row
            ws.Cells(r, 3).value = "TOTAL": ws.Cells(r, 3).Font.Bold = True: ws.Cells(r, 3).HorizontalAlignment = xlRight
            ws.Cells(r, 4).value = totalSafeQty: ws.Cells(r, 4).Font.Bold = True
            ws.Cells(r, 5).value = totalSafe: ws.Cells(r, 5).Font.Bold = True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).HorizontalAlignment = xlLeft
            r = r + 2
        End If
    End If


    ' === MAterials ===
    If includeMat Then
        Set lo = GetTable("tblMaterials")
        If Not lo Is Nothing Then
            ws.Cells(r, 1).value = "Materials": ws.Cells(r, 1).Font.Bold = True: r = r + 1
            Dim hdrMat As Variant: hdrMat = Array("Date", "CategoryID", "ItemDescription", "Quantity", "TotalCost")
            For c = 0 To UBound(hdrMat)
                ws.Cells(r, c + 1).value = hdrMat(c)
                ws.Cells(r, c + 1).Font.Bold = True
                ws.Cells(r, c + 1).HorizontalAlignment = xlCenter
            Next c
            r = r + 1
            totalLog = 0
            For Each cRow In lo.DataBodyRange.rows
                If cRow.Cells(1, ColIndex(lo, "ProjectID")).value = projectID Then
                    Dim mDate As Variant: mDate = cRow.Cells(1, ColIndex(lo, "Date")).value
                    If PassesDateFilter(mDate, dtFrom, dtTo) Then
                        ws.Cells(r, 1).value = mDate
                        ws.Cells(r, 2).value = cRow.Cells(1, ColIndex(lo, "CategoryID")).value
                        ws.Cells(r, 3).value = cRow.Cells(1, ColIndex(lo, "ItemDescription")).value
                        ws.Cells(r, 4).value = cRow.Cells(1, ColIndex(lo, "Quantity")).value
                        ws.Cells(r, 5).value = cRow.Cells(1, ColIndex(lo, "TotalCost")).value
                        
                        ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).HorizontalAlignment = xlLeft
                        ws.Cells(r, 4).HorizontalAlignment = xlLeft
                        ws.Cells(r, 5).HorizontalAlignment = xlLeft
                        
                        totalMatQty = totalMatQty + ws.Cells(r, 4).value
                        totalMat = totalMat + ws.Cells(r, 5).value
                        r = r + 1
                    End If
                End If
            Next cRow
            ' Totals row
            ws.Cells(r, 3).value = "TOTAL": ws.Cells(r, 3).Font.Bold = True: ws.Cells(r, 3).HorizontalAlignment = xlRight
            ws.Cells(r, 4).value = totalMatQty: ws.Cells(r, 4).Font.Bold = True
            ws.Cells(r, 5).value = totalMat: ws.Cells(r, 5).Font.Bold = True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).HorizontalAlignment = xlLeft
            r = r + 2
        End If
    End If


Done:
    ws.Columns.AutoFit
    If wasVisible = xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
    Exit Sub

ErrHandler:
    MsgBox "Error generating report: " & Err.Description, vbCritical
    If ws Is Nothing = False Then
        If wasVisible = xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
    End If
End Sub



' --- helper: find a row in a table by numeric ID column name (returns Range for the matching DataBodyRow) ---
Public Function FindRowByID(lo As ListObject, idColName As String, idVal As Long) As Range
    On Error Resume Next
    Dim c As Range
    If lo Is Nothing Then Exit Function
    Dim idx As Long: idx = ColIndex(lo, idColName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each c In lo.ListColumns(idx).DataBodyRange.Cells
        If val(c.value) = idVal Then
            Set FindRowByID = c.EntireRow
            Exit Function
        End If
    Next c
End Function

Public Sub ExportProjectReportToPDF(Optional ByVal defaultName As String = "")
    Dim ws As Worksheet
    Dim wasVisible As XlSheetVisibility
    Dim fname As Variant
    Dim origPrintArea As String

    Set ws = ThisWorkbook.Worksheets("Rpt_Project")
    wasVisible = ws.Visible

    ' Make visible temporarily for export
    ws.Visible = xlSheetVisible

    ' Save current print area to restore later
    origPrintArea = ws.PageSetup.PrintArea

    ' Ensure the sheet has a defined print area (all used range)
    ws.PageSetup.PrintArea = ws.UsedRange.Address

    ' Set basic page setup
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.45)
        .RightMargin = Application.InchesToPoints(0.45)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With

    ' Prompt user for filename
    fname = Application.GetSaveAsFilename( _
        InitialFileName:=IIf(defaultName = "", "ProjectReport.pdf", defaultName), _
        FileFilter:="PDF Files (*.pdf), *.pdf")

    If fname = False Then GoTo Cleanup

    On Error GoTo ErrHandler
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fname, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    MsgBox "Report exported to: " & fname, vbInformation

Cleanup:
    ' Restore original visibility and print area
    ws.PageSetup.PrintArea = origPrintArea
    If wasVisible = xlSheetVeryHidden Then
        ws.Visible = xlSheetVeryHidden
    Else
        ws.Visible = wasVisible
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error exporting PDF: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub


' --- helper: checks date filter ---
Public Function PassesDateFilter(dVal As Variant, dtFrom As Variant, dtTo As Variant) As Boolean
    ' Default = TRUE (include everything)
    PassesDateFilter = True
    
    ' Skip if no valid date
    If IsEmpty(dVal) Or Not IsDate(dVal) Then Exit Function
    
    ' From-date filter
    If Not IsEmpty(dtFrom) Then
        If dVal < dtFrom Then
            PassesDateFilter = False
            Exit Function
        End If
    End If
    
    ' To-date filter
    If Not IsEmpty(dtTo) Then
        If dVal > dtTo Then
            PassesDateFilter = False
            Exit Function
        End If
    End If
End Function


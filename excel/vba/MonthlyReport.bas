Attribute VB_Name = "MonthlyReport"
'==========================================================================
' Monthly sales-ops pack automation for sales_ops_monthly_report.xlsx
'
' Install: open the workbook > Alt+F11 > File > Import File... > this .bas
'          then File > Save As > Excel Macro-Enabled Workbook (*.xlsm).
'
' Macros
'   BuildMonthlyPack   - for the month in Report!C3: rebuild the regional
'                        pivot, freeze a values-only snapshot sheet, and
'                        export the Report sheet to PDF next to the workbook.
'   BuildAllCompleteMonths - runs BuildMonthlyPack for every complete month.
'==========================================================================
Option Explicit

Private Const REPORT_SHEET As String = "Report"
Private Const MONTH_CELL As String = "C3"
Private Const PIVOT_SHEET As String = "Pivot_Region"
Private Const PIVOT_NAME As String = "pvtRegionMonth"

Public Sub BuildMonthlyPack()
    Dim ws As Worksheet, monthLabel As String
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)
    monthLabel = CStr(ws.Range(MONTH_CELL).Value)
    If Len(monthLabel) = 0 Then
        MsgBox "Pick a report month in " & REPORT_SHEET & "!" & MONTH_CELL & " first.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculate

    BuildRegionPivot
    CreateSnapshot ws, monthLabel
    ExportPdf ws, monthLabel

    ws.Activate
    Application.ScreenUpdating = True
    MsgBox "Monthly pack for " & monthLabel & " is ready.", vbInformation
End Sub

Public Sub BuildAllCompleteMonths()
    Dim lo As ListObject, i As Long, ws As Worksheet, original As String
    Set lo = ThisWorkbook.Worksheets("Data_Monthly").ListObjects("tblMonthly")
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)
    original = CStr(ws.Range(MONTH_CELL).Value)

    Application.DisplayAlerts = False
    For i = 1 To lo.ListRows.Count
        If lo.ListColumns("is_partial_month").DataBodyRange.Cells(i, 1).Value = 0 Then
            ws.Range(MONTH_CELL).Value = lo.ListColumns("month_label").DataBodyRange.Cells(i, 1).Value
            Application.Calculate
            CreateSnapshot ws, CStr(ws.Range(MONTH_CELL).Value)
            ExportPdf ws, CStr(ws.Range(MONTH_CELL).Value)
        End If
    Next i
    Application.DisplayAlerts = True

    ws.Range(MONTH_CELL).Value = original
    Application.Calculate
    MsgBox "Snapshots and PDFs created for all complete months.", vbInformation
End Sub

'--- builds (or rebuilds) a pivot: rows = region, columns = month, values = net sales
Private Sub BuildRegionPivot()
    Dim pws As Worksheet, pc As PivotCache, pt As PivotTable

    On Error Resume Next
    Set pws = ThisWorkbook.Worksheets(PIVOT_SHEET)
    On Error GoTo 0
    If pws Is Nothing Then
        Set pws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        pws.Name = PIVOT_SHEET
    Else
        pws.Cells.Clear
    End If

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="tblRegion")
    Set pt = pc.CreatePivotTable(TableDestination:=pws.Range("A3"), TableName:=PIVOT_NAME)

    With pt
        .PivotFields("region").Orientation = xlRowField
        .PivotFields("month_label").Orientation = xlColumnField
        With .AddDataField(.PivotFields("net_sales"), "Net sales", xlSum)
            .NumberFormat = "£#,##0"
        End With
        .RowAxisLayout xlTabularRow
        .ColumnGrand = True
        .RowGrand = True
    End With
    pws.Range("A1").Value = "Net sales by region and month (pivot rebuilt " & Format(Now, "yyyy-mm-dd hh:nn") & ")"
    pws.Range("A1").Font.Bold = True
End Sub

'--- copies the Report sheet as values only so the month's numbers are frozen
Private Sub CreateSnapshot(src As Worksheet, monthLabel As String)
    Dim snapName As String, snap As Worksheet
    snapName = "Snapshot_" & monthLabel

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(snapName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    src.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set snap = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    snap.Name = snapName
    snap.UsedRange.Value = snap.UsedRange.Value          ' formulas -> values
    snap.Range("B2").Value = "Snapshot taken " & Format(Now, "yyyy-mm-dd hh:nn")
    On Error Resume Next
    snap.Range(MONTH_CELL).Validation.Delete
    On Error GoTo 0
End Sub

Private Sub ExportPdf(ws As Worksheet, monthLabel As String)
    Dim path As String
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the workbook first so the PDF has a folder to go to.", vbExclamation
        Exit Sub
    End If
    path = ThisWorkbook.Path & Application.PathSeparator & "SalesOps_" & monthLabel & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

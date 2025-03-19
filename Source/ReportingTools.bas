Attribute VB_Name = "ReportingTools"
Option Explicit

'---------------------------------------------------------------
' Module: ReportingTools
' Purpose: Automates report generation and formatting
'---------------------------------------------------------------

' GenerateMonthlyReport - Creates monthly summary report from data
Public Sub GenerateMonthlyReport()
    Dim dataSheet As Worksheet
    Dim reportSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim reportDate As Date
    
    ' Set references
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Worksheets("Data")
    If Err.Number <> 0 Then
        MsgBox "Data sheet not found. Please ensure you have a sheet named 'Data'.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Check if report sheet exists, create if not
    On Error Resume Next
    Set reportSheet = ThisWorkbook.Worksheets("Monthly_Report")
    If Err.Number <> 0 Then
        Set reportSheet = ThisWorkbook.Worksheets.Add
        reportSheet.Name = "Monthly_Report"
    Else
        reportSheet.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Determine data range
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row
    lastCol = dataSheet.Cells(1, dataSheet.Columns.Count).End(xlToLeft).Column
    
    ' Create report header
    reportDate = Date
    With reportSheet
        .Range("A1").Value = "Monthly Summary Report"
        .Range("A2").Value = "Generated: " & Format(reportDate, "mmmm d, yyyy")
        .Range("A1:A2").Font.Bold = True
        .Range("A1").Font.Size = 14
    End With
    
    ' Create pivot table for report
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataSheet.Range(dataSheet.Cells(1, 1), dataSheet.Cells(lastRow, lastCol)))
    
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=reportSheet.Range("A5"), _
        TableName:="MonthlySummaryPivot")
    
    ' Configure pivot table fields - adjust these based on your data structure
    With pivotTable
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Category").Position = 1
        
        On Error Resume Next ' In case some fields don't exist
        .PivotFields("Date").Orientation = xlRowField
        .PivotFields("Date").Position = 2
        .PivotFields("Date").Orientation = xlPageField
        
        .PivotFields("Amount").Orientation = xlDataField
        .PivotFields("Amount").Position = 1
        .PivotFields("Amount").Function = xlSum
        .PivotFields("Sum of Amount").NumberFormat = "$#,##0.00"
        
        .PivotFields("Quantity").Orientation = xlDataField
        .PivotFields("Quantity").Function = xlSum
        On Error GoTo 0
    End With
    
    ' Format report
    reportSheet.Columns("A:E").AutoFit
    
    ' Add chart
    Dim chartObj As ChartObject
    Dim chartHeight As Long
    Dim chartWidth As Long
    
    chartHeight = 250
    chartWidth = 375
    
    Set chartObj = reportSheet.ChartObjects.Add( _
        Left:=reportSheet.Range("G5").Left, _
        Top:=reportSheet.Range("G5").Top, _
        Width:=chartWidth, _
        Height:=chartHeight)
    
    With chartObj.Chart
        .SetSourceData Source:=pivotTable.TableRange1
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Monthly Summary by Category"
        .Legend.Position = xlLegendPositionBottom
    End With
    
    MsgBox "Monthly report has been generated!", vbInformation
End Sub

' ExportAsPDF - Exports current sheet as PDF
Public Sub ExportAsPDF()
    Dim pdfPath As String
    Dim ws As Worksheet
    
    ' Get active sheet
    Set ws = ActiveSheet
    
    ' Set default file name based on sheet name
    pdfPath = ThisWorkbook.Path & "\" & ws.Name & "_" & Format(Now(), "yyyymmdd") & ".pdf"
    
    ' Ask for file location
    pdfPath = Application.GetSaveAsFilename( _
        InitialFileName:=pdfPath, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Report as PDF")
    
    If pdfPath = "False" Then Exit Sub ' User canceled
    
    ' Export as PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    MsgBox "PDF export complete!", vbInformation
End Sub

' ScheduleReports - Sets up automatic report generation schedule
Public Sub ScheduleReports()
    ' This is a placeholder function as VBA cannot directly schedule tasks
    ' In a real implementation, this would either:
    ' 1. Create a Windows Task Scheduler task via shell
    ' 2. Set up internal timing mechanism for when Excel is open
    ' 3. Provide instructions for manual scheduling
    
    MsgBox "To schedule this report:" & vbCrLf & _
           "1. Open Windows Task Scheduler" & vbCrLf & _
           "2. Create a new Basic Task" & vbCrLf & _
           "3. Set your preferred schedule" & vbCrLf & _
           "4. Action: Start a Program" & vbCrLf & _
           "5. Program: Excel.exe" & vbCrLf & _
           "6. Add arguments: /e " & Chr(34) & ThisWorkbook.FullName & Chr(34) & " /x GenerateMonthlyReport", _
           vbInformation, "Report Scheduling"
End Sub 
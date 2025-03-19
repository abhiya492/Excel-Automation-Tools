Attribute VB_Name = "DataProcessing"
Option Explicit

'---------------------------------------------------------------
' Module: DataProcessing
' Purpose: Contains functions for automated data processing tasks
'---------------------------------------------------------------

' CleanData - Removes duplicates, trims whitespace, and standardizes formats
Public Sub CleanData(targetRange As Range)
    Application.ScreenUpdating = False
    
    ' Remove duplicates
    targetRange.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
    
    ' Trim whitespace and standardize case for text cells
    Dim cell As Range
    For Each cell In targetRange
        If Not IsEmpty(cell) And IsNumeric(cell.Value) = False Then
            cell.Value = WorksheetFunction.Trim(cell.Value)
        End If
    Next cell
    
    Application.ScreenUpdating = True
    
    MsgBox "Data cleaning complete!", vbInformation
End Sub

' ImportCSVData - Imports and formats data from CSV files
Public Sub ImportCSVData()
    Dim filePath As Variant
    Dim targetSheet As Worksheet
    
    ' Prompt user to select CSV file
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV file to import")
    
    If filePath = False Then Exit Sub ' User canceled
    
    ' Create new worksheet for imported data
    Set targetSheet = ThisWorkbook.Worksheets.Add
    targetSheet.Name = "Imported_" & Format(Now(), "yyyymmdd")
    
    ' Import the CSV data
    With targetSheet.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=targetSheet.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
    
    ' Auto-fit columns
    targetSheet.Cells.EntireColumn.AutoFit
    
    MsgBox "Data imported successfully!", vbInformation
End Sub

' BatchProcessFiles - Process multiple Excel files from a folder
Public Sub BatchProcessFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim summarySheet As Worksheet
    Dim lastRow As Long
    
    ' Prompt user for folder path
    folderPath = BrowseForFolder()
    If folderPath = "" Then Exit Sub
    
    ' Create summary sheet
    Set summarySheet = ThisWorkbook.Worksheets.Add
    summarySheet.Name = "Batch_Summary"
    
    ' Add headers
    With summarySheet
        .Range("A1").Value = "File Name"
        .Range("B1").Value = "Record Count"
        .Range("C1").Value = "Date Processed"
        .Range("D1").Value = "Status"
        .Range("A1:D1").Font.Bold = True
    End With
    
    ' Get first Excel file in the folder
    fileName = Dir(folderPath & "*.xls*")
    lastRow = 2
    
    Application.ScreenUpdating = False
    
    ' Loop through all Excel files in the folder
    Do While fileName <> ""
        ' Add to summary
        summarySheet.Range("A" & lastRow).Value = fileName
        summarySheet.Range("C" & lastRow).Value = Now()
        
        On Error Resume Next
        Set wb = Workbooks.Open(folderPath & fileName)
        
        If Err.Number = 0 Then
            ' Process the file - example just counts rows in first sheet
            summarySheet.Range("B" & lastRow).Value = wb.Worksheets(1).UsedRange.Rows.Count
            summarySheet.Range("D" & lastRow).Value = "Processed"
            
            wb.Close SaveChanges:=False
        Else
            summarySheet.Range("D" & lastRow).Value = "Error: " & Err.Description
        End If
        On Error GoTo 0
        
        lastRow = lastRow + 1
        fileName = Dir() ' Get next file
    Loop
    
    Application.ScreenUpdating = True
    summarySheet.Columns("A:D").AutoFit
    
    MsgBox "Batch processing complete!", vbInformation
End Sub

' Helper function to browse for folder
Private Function BrowseForFolder() As String
    Dim folderDialog As Object
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .Title = "Select Folder Containing Excel Files"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1) & "\"
        Else
            BrowseForFolder = ""
        End If
    End With
End Function 
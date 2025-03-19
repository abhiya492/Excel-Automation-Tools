Attribute VB_Name = "UserInterface"
Option Explicit

'---------------------------------------------------------------
' Module: UserInterface
' Purpose: Interface elements and form controls for non-technical users
'---------------------------------------------------------------

' CreateMainMenu - Creates a main menu interface with buttons on a new sheet
Public Sub CreateMainMenu()
    Dim menuSheet As Worksheet
    Dim btn As Button
    Dim buttonTop As Long
    Dim buttonLeft As Long
    Dim buttonWidth As Long
    Dim buttonHeight As Long
    Dim i As Integer
    
    ' Check if menu sheet exists and create if not
    On Error Resume Next
    Set menuSheet = ThisWorkbook.Worksheets("Main_Menu")
    If Err.Number <> 0 Then
        Set menuSheet = ThisWorkbook.Worksheets.Add
        menuSheet.Name = "Main_Menu"
    Else
        ' Clear existing buttons
        For i = menuSheet.Buttons.Count To 1 Step -1
            menuSheet.Buttons(i).Delete
        Next i
    End If
    On Error GoTo 0
    
    ' Format menu sheet
    With menuSheet
        .Range("A:Z").ColumnWidth = 15
        .Range("A:Z").HorizontalAlignment = xlCenter
        .Cells.ClearContents
        
        ' Add title
        .Range("B2:G2").Merge
        .Range("B2").Value = "Excel Automation Toolkit - Main Menu"
        .Range("B2").Font.Size = 16
        .Range("B2").Font.Bold = True
        
        ' Add subtitle
        .Range("B3:G3").Merge
        .Range("B3").Value = "Select a function below to begin"
        .Range("B3").Font.Italic = True
        
        ' Add sections
        .Range("B5").Value = "Data Processing"
        .Range("B5").Font.Bold = True
        .Range("B5").Font.Size = 12
        
        .Range("D5").Value = "Reporting"
        .Range("D5").Font.Bold = True
        .Range("D5").Font.Size = 12
        
        .Range("F5").Value = "Utilities"
        .Range("F5").Font.Bold = True
        .Range("F5").Font.Size = 12
    End With
    
    ' Button dimensions
    buttonWidth = 120
    buttonHeight = 30
    
    ' Data Processing Section
    buttonLeft = menuSheet.Range("B6").Left
    buttonTop = menuSheet.Range("B6").Top
    
    ' Import Data button
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Import Data"
        .Name = "btnImportData"
        .OnAction = "DataProcessing.ImportCSVData"
    End With
    
    ' Clean Data button
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Clean Data"
        .Name = "btnCleanData"
        .OnAction = "ShowCleanDataForm"
    End With
    
    ' Batch Process button
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Batch Process Files"
        .Name = "btnBatchProcess"
        .OnAction = "DataProcessing.BatchProcessFiles"
    End With
    
    ' Reporting Section
    buttonLeft = menuSheet.Range("D6").Left
    buttonTop = menuSheet.Range("D6").Top
    
    ' Generate Report button
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Generate Report"
        .Name = "btnGenerateReport"
        .OnAction = "ReportingTools.GenerateMonthlyReport"
    End With
    
    ' Export PDF button
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Export as PDF"
        .Name = "btnExportPDF"
        .OnAction = "ReportingTools.ExportAsPDF"
    End With
    
    ' Schedule Reports button
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Schedule Reports"
        .Name = "btnScheduleReports"
        .OnAction = "ReportingTools.ScheduleReports"
    End With
    
    ' Utilities Section
    buttonLeft = menuSheet.Range("F6").Left
    buttonTop = menuSheet.Range("F6").Top
    
    ' Validate Data button
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Validate Data"
        .Name = "btnValidateData"
        .OnAction = "DataValidation.ValidateDataSheet"
    End With
    
    ' Apply Validation button
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Apply Validation Rules"
        .Name = "btnApplyValidation"
        .OnAction = "DataValidation.ApplyDataValidation"
    End With
    
    ' Help button
    buttonTop = buttonTop + buttonHeight + 5
    Set btn = menuSheet.Buttons.Add(buttonLeft, buttonTop, buttonWidth, buttonHeight)
    With btn
        .Caption = "Help"
        .Name = "btnHelp"
        .OnAction = "ShowHelpForm"
    End With
    
    ' Add a note at the bottom
    menuSheet.Range("B15:F15").Merge
    menuSheet.Range("B15").Value = "Note: Some functions require specific data formats or sheet structures."
    menuSheet.Range("B15").Font.Italic = True
    
    ' Activate the menu sheet
    menuSheet.Activate
    
    MsgBox "Main menu created successfully!", vbInformation
End Sub

' ShowCleanDataForm - Form for Clean Data function
' Note: In a real implementation, this would create a UserForm
Public Sub ShowCleanDataForm()
    Dim targetRange As Range
    
    On Error Resume Next
    Set targetRange = Application.InputBox("Select the range to clean:", "Clean Data", Type:=8)
    If Err.Number <> 0 Or targetRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Call the cleaning function
    DataProcessing.CleanData targetRange
End Sub

' ShowHelpForm - Shows help information
' Note: In a real implementation, this would create a UserForm
Public Sub ShowHelpForm()
    Dim helpMessage As String
    
    helpMessage = "Excel Automation Toolkit Help" & vbCrLf & vbCrLf & _
                  "Data Processing:" & vbCrLf & _
                  "- Import Data: Imports data from CSV files" & vbCrLf & _
                  "- Clean Data: Removes duplicates and standardizes formats" & vbCrLf & _
                  "- Batch Process: Process multiple Excel files at once" & vbCrLf & vbCrLf & _
                  "Reporting:" & vbCrLf & _
                  "- Generate Report: Creates monthly summary report" & vbCrLf & _
                  "- Export as PDF: Saves current sheet as PDF" & vbCrLf & _
                  "- Schedule Reports: Set up automated report generation" & vbCrLf & vbCrLf & _
                  "Utilities:" & vbCrLf & _
                  "- Validate Data: Checks data against validation rules" & vbCrLf & _
                  "- Apply Validation: Adds Excel data validation to cells" & vbCrLf & vbCrLf & _
                  "For more detailed help, please refer to the documentation."
    
    MsgBox helpMessage, vbInformation, "Excel Automation Toolkit Help"
End Sub

' CreateDataEntryForm - Creates a data entry form with validation
' Note: This would typically be implemented as a UserForm in VBA
Public Sub CreateDataEntryForm()
    Dim dataSheet As Worksheet
    Dim formSheet As Worksheet
    Dim buttonTop As Long
    Dim buttonLeft As Long
    Dim btn As Button
    
    ' Check if form sheet exists and create if not
    On Error Resume Next
    Set formSheet = ThisWorkbook.Worksheets("Data_Entry_Form")
    If Err.Number <> 0 Then
        Set formSheet = ThisWorkbook.Worksheets.Add
        formSheet.Name = "Data_Entry_Form"
    Else
        formSheet.Cells.Clear
        ' Clear existing form controls
        On Error Resume Next
        For Each btn In formSheet.Buttons
            btn.Delete
        Next btn
        On Error GoTo 0
    End If
    On Error GoTo 0
    
    ' Check if data sheet exists and create if not
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Worksheets("Data")
    If Err.Number <> 0 Then
        Set dataSheet = ThisWorkbook.Worksheets.Add
        dataSheet.Name = "Data"
        
        ' Initialize data sheet with headers
        With dataSheet
            .Range("A1").Value = "ID"
            .Range("B1").Value = "Date"
            .Range("C1").Value = "Category"
            .Range("D1").Value = "Description"
            .Range("E1").Value = "Amount"
            .Range("F1").Value = "Status"
            
            .Range("A1:F1").Font.Bold = True
        End With
    End If
    On Error GoTo 0
    
    ' Format form sheet
    With formSheet
        .Range("A:Z").ColumnWidth = 15
        
        ' Add title
        .Range("B2:E2").Merge
        .Range("B2").Value = "Data Entry Form"
        .Range("B2").Font.Size = 14
        .Range("B2").Font.Bold = True
        .Range("B2").HorizontalAlignment = xlCenter
        
        ' Add form labels and input cells
        .Range("B4").Value = "ID:"
        .Range("B4").Font.Bold = True
        .Range("C4").Value = "AUTO"
        .Range("C4").Interior.Color = RGB(220, 220, 220)
        
        .Range("B5").Value = "Date:"
        .Range("B5").Font.Bold = True
        .Range("C5").NumberFormat = "mm/dd/yyyy"
        .Range("C5").Value = Date
        
        .Range("B6").Value = "Category:"
        .Range("B6").Font.Bold = True
        
        ' Add data validation for category
        With .Range("C6").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="Sales,Expenses,Inventory,Marketing,Other"
            .InputTitle = "Select Category"
            .InputMessage = "Please select a category from the list"
            .ErrorTitle = "Invalid Category"
            .ErrorMessage = "Please select a category from the dropdown list"
            .ShowInput = True
            .ShowError = True
        End With
        
        .Range("B7").Value = "Description:"
        .Range("B7").Font.Bold = True
        
        .Range("B8").Value = "Amount:"
        .Range("B8").Font.Bold = True
        .Range("C8").NumberFormat = "$#,##0.00"
        
        .Range("B9").Value = "Status:"
        .Range("B9").Font.Bold = True
        
        ' Add data validation for status
        With .Range("C9").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="Pending,Completed,Cancelled"
            .InputTitle = "Select Status"
            .InputMessage = "Please select a status from the list"
            .ErrorTitle = "Invalid Status"
            .ErrorMessage = "Please select a status from the dropdown list"
            .ShowInput = True
            .ShowError = True
        End With
    End With
    
    ' Add buttons
    buttonLeft = formSheet.Range("B11").Left
    buttonTop = formSheet.Range("B11").Top
    
    ' Submit button
    Set btn = formSheet.Buttons.Add(buttonLeft, buttonTop, 75, 25)
    With btn
        .Caption = "Submit"
        .Name = "btnSubmit"
        .OnAction = "SubmitDataForm"
    End With
    
    ' Clear button
    buttonLeft = formSheet.Range("C11").Left
    Set btn = formSheet.Buttons.Add(buttonLeft, buttonTop, 75, 25)
    With btn
        .Caption = "Clear"
        .Name = "btnClearForm"
        .OnAction = "ClearDataForm"
    End With
    
    ' Close button
    buttonLeft = formSheet.Range("D11").Left
    Set btn = formSheet.Buttons.Add(buttonLeft, buttonTop, 75, 25)
    With btn
        .Caption = "Close"
        .Name = "btnCloseForm"
        .OnAction = "CloseDataForm"
    End With
    
    ' Add instructions
    formSheet.Range("B13:E15").Merge
    formSheet.Range("B13").Value = "Instructions: Fill in all fields and click Submit to add data. " & _
                                 "The ID will be assigned automatically. Required fields are in bold."
    formSheet.Range("B13").Font.Italic = True
    formSheet.Range("B13").WrapText = True
    
    ' Activate the form sheet
    formSheet.Activate
    
    MsgBox "Data entry form created successfully!", vbInformation
End Sub

' SubmitDataForm - Processes the data entry form submission
' Note: In a real implementation, this would be part of a UserForm
Public Sub SubmitDataForm()
    Dim formSheet As Worksheet
    Dim dataSheet As Worksheet
    Dim lastRow As Long
    Dim newID As String
    
    ' Get sheet references
    Set formSheet = ThisWorkbook.Worksheets("Data_Entry_Form")
    Set dataSheet = ThisWorkbook.Worksheets("Data")
    
    ' Validate form data
    If formSheet.Range("C6").Value = "" Then
        MsgBox "Please select a Category.", vbExclamation
        Exit Sub
    End If
    
    If formSheet.Range("C7").Value = "" Then
        MsgBox "Description is required.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(formSheet.Range("C8").Value) Then
        MsgBox "Amount must be a numeric value.", vbExclamation
        Exit Sub
    End If
    
    If formSheet.Range("C9").Value = "" Then
        MsgBox "Please select a Status.", vbExclamation
        Exit Sub
    End If
    
    ' Get last row in data sheet and generate ID
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row + 1
    newID = "ID" & Format(lastRow - 1, "000")
    
    ' Add data to data sheet
    With dataSheet
        .Cells(lastRow, 1).Value = newID
        .Cells(lastRow, 2).Value = formSheet.Range("C5").Value
        .Cells(lastRow, 3).Value = formSheet.Range("C6").Value
        .Cells(lastRow, 4).Value = formSheet.Range("C7").Value
        .Cells(lastRow, 5).Value = formSheet.Range("C8").Value
        .Cells(lastRow, 6).Value = formSheet.Range("C9").Value
    End With
    
    ' Clear form for next entry
    ClearDataForm
    
    ' Update ID display for next entry
    formSheet.Range("C4").Value = "ID" & Format(lastRow, "000")
    
    MsgBox "Data added successfully!", vbInformation
End Sub

' ClearDataForm - Clears the data entry form
Public Sub ClearDataForm()
    Dim formSheet As Worksheet
    Set formSheet = ThisWorkbook.Worksheets("Data_Entry_Form")
    
    With formSheet
        .Range("C5").Value = Date
        .Range("C6").Value = ""
        .Range("C7").Value = ""
        .Range("C8").Value = ""
        .Range("C9").Value = ""
    End With
End Sub

' CloseDataForm - Returns to main menu
Public Sub CloseDataForm()
    ThisWorkbook.Worksheets("Main_Menu").Activate
End Sub 
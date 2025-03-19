Attribute VB_Name = "DataValidation"
Option Explicit

'---------------------------------------------------------------
' Module: DataValidation
' Purpose: Data validation and error checking routines
'---------------------------------------------------------------

' ValidateDataSheet - Performs comprehensive data validation on a worksheet
Public Sub ValidateDataSheet(Optional sheetName As String = "")
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerRow As Long
    Dim errorCount As Long
    Dim errorSheet As Worksheet
    Dim cell As Range
    Dim errorLog As String
    Dim i As Long, j As Long
    Dim validationRules As Collection
    Dim dataTypes As Collection
    
    ' Set reference to worksheet
    If sheetName = "" Then
        Set ws = ActiveSheet
    Else
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetName)
        If Err.Number <> 0 Then
            MsgBox "Sheet '" & sheetName & "' not found!", vbExclamation
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    ' Create error log sheet if it doesn't exist
    On Error Resume Next
    Set errorSheet = ThisWorkbook.Worksheets("ValidationErrors")
    If Err.Number <> 0 Then
        Set errorSheet = ThisWorkbook.Worksheets.Add
        errorSheet.Name = "ValidationErrors"
    Else
        errorSheet.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Set up error log sheet
    With errorSheet
        .Range("A1").Value = "Validation Report"
        .Range("A2").Value = "Date: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        .Range("A3").Value = "Sheet: " & ws.Name
        .Range("A1:A3").Font.Bold = True
        
        .Range("A5").Value = "Cell"
        .Range("B5").Value = "Value"
        .Range("C5").Value = "Expected Type"
        .Range("D5").Value = "Error Description"
        .Range("A5:D5").Font.Bold = True
    End With
    
    ' Determine data range
    headerRow = 1 ' Assuming first row is header
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' Create collection of validation rules based on headers
    Set validationRules = New Collection
    Set dataTypes = New Collection
    
    ' Example validation rules - in real application, these would be customizable
    ' or loaded from a configuration
    For j = 1 To lastCol
        Select Case ws.Cells(headerRow, j).Value
            Case "ID", "EmployeeID", "CustomerID"
                validationRules.Add "^[A-Z0-9]{5,10}$" ' Alphanumeric ID format
                dataTypes.Add "ID"
            Case "Email"
                validationRules.Add "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$" ' Email format
                dataTypes.Add "Email"
            Case "Phone"
                validationRules.Add "^\d{3}-\d{3}-\d{4}$" ' Phone format (e.g., 555-123-4567)
                dataTypes.Add "Phone"
            Case "Date", "StartDate", "EndDate", "DOB"
                validationRules.Add "" ' Date validation handled separately
                dataTypes.Add "Date" 
            Case "Amount", "Price", "Cost", "Salary"
                validationRules.Add "" ' Numeric validation handled separately
                dataTypes.Add "Currency"
            Case "Quantity", "Count", "Number"
                validationRules.Add "" ' Integer validation handled separately
                dataTypes.Add "Integer"
            Case "Percentage", "Rate"
                validationRules.Add "" ' Percentage validation handled separately
                dataTypes.Add "Percentage"
            Case Else
                validationRules.Add "" ' No specific validation
                dataTypes.Add "Text"
        End Select
    Next j
    
    ' Initialize error count
    errorCount = 0
    
    ' Validate each cell against its rule
    For i = headerRow + 1 To lastRow
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            If Not IsEmpty(cell.Value) Then
                Select Case dataTypes(j)
                    Case "ID", "Email", "Phone"
                        ' Regex pattern validation
                        If Not RegExTest(CStr(cell.Value), validationRules(j)) Then
                            LogError errorSheet, errorCount, cell, dataTypes(j), _
                                "Invalid format for " & dataTypes(j)
                        End If
                    
                    Case "Date"
                        ' Date validation
                        If Not IsDate(cell.Value) Then
                            LogError errorSheet, errorCount, cell, "Date", "Invalid date format"
                        End If
                    
                    Case "Currency"
                        ' Currency validation (numeric with optional decimal)
                        If Not IsNumeric(cell.Value) Then
                            LogError errorSheet, errorCount, cell, "Currency", "Value must be numeric"
                        End If
                    
                    Case "Integer"
                        ' Integer validation
                        If Not IsNumeric(cell.Value) Or Int(cell.Value) <> cell.Value Then
                            LogError errorSheet, errorCount, cell, "Integer", "Value must be an integer"
                        End If
                    
                    Case "Percentage"
                        ' Percentage validation (numeric between 0-100)
                        If Not IsNumeric(cell.Value) Or cell.Value < 0 Or cell.Value > 100 Then
                            LogError errorSheet, errorCount, cell, "Percentage", _
                                "Value must be a percentage (0-100)"
                        End If
                End Select
            End If
        Next j
    Next i
    
    ' Format error log sheet
    errorSheet.Columns("A:D").AutoFit
    
    ' Highlight the worksheet cells with errors
    If errorCount > 0 Then
        MsgBox "Validation complete. " & errorCount & " errors found. " & _
               "See 'ValidationErrors' sheet for details.", vbExclamation
    Else
        MsgBox "Validation complete. No errors found!", vbInformation
    End If
End Sub

' LogError - Helper function to log validation errors
Private Sub LogError(errorSheet As Worksheet, ByRef errorCount As Long, _
                    cell As Range, expectedType As String, errorDesc As String)
    errorCount = errorCount + 1
    
    With errorSheet
        .Cells(errorCount + 5, 1).Value = cell.Address(False, False)
        .Cells(errorCount + 5, 2).Value = cell.Value
        .Cells(errorCount + 5, 3).Value = expectedType
        .Cells(errorCount + 5, 4).Value = errorDesc
    End With
    
    ' Mark the cell with error in the original sheet
    cell.Interior.Color = RGB(255, 200, 200)
End Sub

' RegExTest - Helper function for regex validation
' Note: VBA doesn't have built-in regex, this is a simplified implementation
Private Function RegExTest(text As String, pattern As String) As Boolean
    ' Basic implementation of common patterns - in a real app, use VBScript.RegExp
    Select Case pattern
        Case "^[A-Z0-9]{5,10}$" ' ID format
            RegExTest = Len(text) >= 5 And Len(text) <= 10 And _
                       Not (text Like "*[!A-Z0-9]*")
        
        Case "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$" ' Email
            RegExTest = text Like "*@*.?*" And _
                       InStr(text, "@") > 1 And _
                       InStr(InStr(text, "@") + 1, text, ".") > InStr(text, "@") + 1
        
        Case "^\d{3}-\d{3}-\d{4}$" ' Phone format
            RegExTest = Len(text) = 12 And _
                       Mid(text, 4, 1) = "-" And _
                       Mid(text, 8, 1) = "-" And _
                       IsNumeric(Left(text, 3)) And _
                       IsNumeric(Mid(text, 5, 3)) And _
                       IsNumeric(Right(text, 4))
        
        Case Else
            RegExTest = True ' No validation
    End Select
End Function

' ApplyDataValidation - Applies Excel data validation to a range
Public Sub ApplyDataValidation()
    Dim targetRange As Range
    Dim validationType As Integer
    Dim title As String
    Dim prompt As String
    Dim errorTitle As String
    Dim errorMsg As String
    
    ' Ask user to select range
    On Error Resume Next
    Set targetRange = Application.InputBox("Select the range to apply validation to:", "Data Validation", Type:=8)
    If Err.Number <> 0 Or targetRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    ' Show validation type form
    validationType = ShowValidationForm()
    
    If validationType = 0 Then Exit Sub ' User canceled
    
    ' Common validation messages
    title = "Data Validation"
    prompt = "Please enter a value that meets the validation criteria."
    errorTitle = "Invalid Entry"
    errorMsg = "The value you entered does not meet the validation criteria."
    
    ' Apply validation based on type
    With targetRange.Validation
        .Delete
        
        Select Case validationType
            Case 1 ' Text Length
                .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, Formula1:="1", Formula2:="255"
            
            Case 2 ' Whole Number
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlGreaterEqual, Formula1:="0"
            
            Case 3 ' Decimal
                .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlGreaterEqual, Formula1:="0"
            
            Case 4 ' Date
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlGreaterEqual, Formula1:="1/1/2000"
            
            Case 5 ' List
                Dim listItems As String
                listItems = Application.InputBox("Enter comma-separated list items:", "List Validation")
                If listItems = "" Or listItems = "False" Then Exit Sub
                
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Formula1:=listItems
            
            Case 6 ' Custom formula
                Dim formula As String
                formula = Application.InputBox("Enter validation formula (e.g., =A1>0):", "Custom Validation")
                If formula = "" Or formula = "False" Then Exit Sub
                
                If Left(formula, 1) <> "=" Then formula = "=" & formula
                
                .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
                    Formula1:=formula
        End Select
        
        ' Set validation messages
        .InputTitle = title
        .InputMessage = prompt
        .ErrorTitle = errorTitle
        .ErrorMessage = errorMsg
        .ShowInput = True
        .ShowError = True
    End With
    
    MsgBox "Data validation applied successfully!", vbInformation
End Sub

' ShowValidationForm - Shows a form to select validation type
' Note: This is a simplified version, in a real app use a UserForm
Private Function ShowValidationForm() As Integer
    Dim msg As String
    
    msg = "Select validation type:" & vbCrLf & vbCrLf & _
          "1. Text Length" & vbCrLf & _
          "2. Whole Number" & vbCrLf & _
          "3. Decimal" & vbCrLf & _
          "4. Date" & vbCrLf & _
          "5. List" & vbCrLf & _
          "6. Custom Formula" & vbCrLf & vbCrLf & _
          "Enter number (0 to cancel):"
    
    Dim result As String
    result = InputBox(msg, "Data Validation Type")
    
    If result = "" Then
        ShowValidationForm = 0 ' Canceled
    Else
        ShowValidationForm = Val(result)
        If ShowValidationForm < 1 Or ShowValidationForm > 6 Then
            ShowValidationForm = 0 ' Invalid entry
        End If
    End If
End Function 
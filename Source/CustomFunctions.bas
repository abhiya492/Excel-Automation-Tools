Attribute VB_Name = "CustomFunctions"
Option Explicit

'---------------------------------------------------------------
' Module: CustomFunctions
' Purpose: Custom worksheet functions to standardize calculations
'---------------------------------------------------------------

' RiskAdjustedValue - Calculates a risk-adjusted value based on input parameters
Public Function RiskAdjustedValue(value As Double, riskFactor As Double) As Double
    ' Calculate risk-adjusted value with validation
    If riskFactor < 0 Or riskFactor > 1 Then
        RiskAdjustedValue = CVErr(xlErrValue)
    Else
        RiskAdjustedValue = value * (1 - riskFactor)
    End If
End Function

' WeightedAverage - Calculates weighted average of values
Public Function WeightedAverage(valuesRange As Range, weightsRange As Range) As Variant
    Dim totalWeight As Double
    Dim weightedSum As Double
    Dim i As Long
    
    ' Validate input ranges
    If valuesRange.Count <> weightsRange.Count Then
        WeightedAverage = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Calculate weighted average
    For i = 1 To valuesRange.Count
        ' Check if values are numeric
        If Not IsNumeric(valuesRange(i).Value) Or Not IsNumeric(weightsRange(i).Value) Then
            WeightedAverage = CVErr(xlErrValue)
            Exit Function
        End If
        
        weightedSum = weightedSum + (valuesRange(i).Value * weightsRange(i).Value)
        totalWeight = totalWeight + weightsRange(i).Value
    Next i
    
    ' Prevent division by zero
    If totalWeight = 0 Then
        WeightedAverage = CVErr(xlErrDiv0)
    Else
        WeightedAverage = weightedSum / totalWeight
    End If
End Function

' FiscalQuarter - Returns fiscal quarter based on date and fiscal year start month
Public Function FiscalQuarter(inputDate As Date, Optional fiscalYearStartMonth As Integer = 1) As Integer
    Dim monthDiff As Integer
    
    ' Validate fiscal year start month
    If fiscalYearStartMonth < 1 Or fiscalYearStartMonth > 12 Then
        FiscalQuarter = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Calculate months since fiscal year start
    monthDiff = (Month(inputDate) - fiscalYearStartMonth) Mod 12
    If monthDiff < 0 Then monthDiff = monthDiff + 12
    
    ' Determine quarter
    FiscalQuarter = (monthDiff \ 3) + 1
End Function

' BusinessDaysDifference - Calculates number of business days between two dates
Public Function BusinessDaysDifference(startDate As Date, endDate As Date, _
                                      Optional includeHolidays As Boolean = False) As Variant
    Dim dayCount As Long
    Dim currDate As Date
    Dim holidays As Variant
    
    ' Validate date order
    If startDate > endDate Then
        BusinessDaysDifference = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Set default holidays list (US holidays as example)
    If includeHolidays Then
        holidays = Array( _
            DateSerial(Year(startDate), 1, 1),    ' New Year's Day
            DateSerial(Year(startDate), 7, 4),    ' Independence Day
            DateSerial(Year(startDate), 12, 25),  ' Christmas
            DateSerial(Year(endDate), 1, 1),      ' New Year's Day
            DateSerial(Year(endDate), 7, 4),      ' Independence Day
            DateSerial(Year(endDate), 12, 25)     ' Christmas
        )
    End If
    
    ' Initialize counter
    dayCount = 0
    currDate = startDate
    
    ' Count business days
    Do While currDate <= endDate
        ' Check if current date is a weekday (not Saturday or Sunday)
        If Weekday(currDate) <> vbSaturday And Weekday(currDate) <> vbSunday Then
            ' Check if it's not a holiday (if checking for holidays)
            If includeHolidays Then
                Dim isHoliday As Boolean
                Dim i As Integer
                
                isHoliday = False
                For i = LBound(holidays) To UBound(holidays)
                    If currDate = holidays(i) Then
                        isHoliday = True
                        Exit For
                    End If
                Next i
                
                If Not isHoliday Then dayCount = dayCount + 1
            Else
                dayCount = dayCount + 1
            End If
        End If
        
        ' Move to next day
        currDate = currDate + 1
    Loop
    
    BusinessDaysDifference = dayCount
End Function

' FormatCurrency - Standardized currency formatting with options
Public Function FormatCurrency(value As Double, Optional currencySymbol As String = "$", _
                               Optional decimalPlaces As Integer = 2, _
                               Optional includeThousandsSeparator As Boolean = True) As String
    Dim formatString As String
    Dim result As String
    
    ' Validate decimal places
    If decimalPlaces < 0 Or decimalPlaces > 10 Then
        FormatCurrency = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Build format string
    If includeThousandsSeparator Then
        formatString = "#,##0"
    Else
        formatString = "0"
    End If
    
    ' Add decimal places if needed
    If decimalPlaces > 0 Then
        formatString = formatString & "." & String(decimalPlaces, "0")
    End If
    
    ' Format the value
    result = Format(value, formatString)
    
    ' Add currency symbol
    FormatCurrency = currencySymbol & result
End Function 
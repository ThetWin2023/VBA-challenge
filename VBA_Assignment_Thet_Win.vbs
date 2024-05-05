' Data Analytics and Visualization Bootcamp
' Module 2 - VBA Scripting
' Version 1.1
' Assignment 2
' Name: Thet Win
' Date: May 5, 2024

Sub main()

    For Each ws In Worksheets
        ws.Activate
        Call TickerSummary
        Call Format
    Next ws

End Sub
Sub TickerSummary()

    'Declare and Initialize the counter variables
    Dim myRow, myCounter, tickerCounter, closeCounter As Integer
    Dim lastRow As Double
    
    'Declare and initialize calculated variables
    Dim openValue, closeValue, quantityChange, percentChange As Integer
    Dim totalVolume As Double
    
    'Declare Summary MAX and MIN percent values
    Dim maxPercent, minPercent, maxVolume As Double
    
    'Variables for Max and Min Percent change summary table
    Dim maxPercentRange, minPercentRange, maxVolRange As Range
    
    'Initialize variables
    myRow = 0
    myCounter = 1
    tickerCounter = 1   'ticker value row counter
    closeCounter = 2    'closingValue row counter
    lastRow = 0
    maxPercent = 0
    minPercent = 0
    maxVolume = 0
    
    

    
    'Get the last row
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Headers
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Quantity Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    Cells(1, 16) = "Ticker"
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    
    'Temporary Table for Calculations
    Cells(1, 18) = "Ticker"
    Cells(1, 19) = "Opening Value"
    Cells(1, 20) = "Closing Value"
    
    
    For myRow = 2 To lastRow
        
        If Cells(myRow, 1).Value = Cells(myCounter, 9).Value Then
            GoTo nxtRow
        Else
        
            'Reset for new Ticker value
            openValue = 0
            closeValue = 0
            quantityChange = 0
            percentChange = 0
            totalVolume = 0
    
            myCounter = myCounter + 1
            Cells(myCounter, 9).Value = Cells(myRow, 1).Value
            
            'Temporary Table for Calculations
            Cells(myCounter, 18).Value = Cells(myRow, 1).Value  'Ticker
            'Opening Value
            Cells(myCounter, 19).Value = Cells(myRow, 3).Value  'Opening Value
            openValue = Cells(myRow, 3).Value

            'Closing Value - Loop through until the next Ticker is Found
            tickerCounter = myRow
            Do While Cells(tickerCounter, 1).Value = Cells(myCounter, 18).Value
                tickerCounter = tickerCounter + 1
                If tickerCounter = lastRow Then
                    Cells(closeCounter, 20).Value = Cells(tickerCounter - 1, 6).Value
                    
                    closeValue = Cells(tickerCounter - 1, 6).Value
                    closeCounter = closeCounter + 1
                    
                    'calculate quantity change, percent change and total stock volume
                    quantityChange = closeValue - openValue
                    totalVolume = Application.Sum(Range(Cells(myRow, 7), Cells(tickerCounter - 1, 7)))
                    
                    Cells(myCounter, 12) = totalVolume
                    Cells(myCounter, 11) = (quantityChange / ((closeValue + openValue) / 2))
                    Cells(myCounter, 10) = quantityChange
                    GoTo getSummary
                End If
            Loop
            Cells(closeCounter, 20).Value = Cells(tickerCounter - 1, 6).Value 'Populate Closing Value once next ticker value is found
            closeValue = Cells(tickerCounter - 1, 6).Value
            closeCounter = closeCounter + 1
            
            'calculate quantity change, percent change and total stock volume
            quantityChange = closeValue - openValue
            totalVolume = Application.Sum(Range(Cells(myRow, 7), Cells(tickerCounter - 1, 7)))
            
            Cells(myCounter, 12) = totalVolume
            Cells(myCounter, 11) = (quantityChange / ((closeValue + openValue) / 2))
            Cells(myCounter, 10) = quantityChange

        End If
        
nxtRow:
    Next myRow
    
getSummary:
    'Get Greatest (MAX) values Summary
    
    maxPercent = Application.WorksheetFunction.Max(Range("K:K"))
    minPercent = Application.WorksheetFunction.Min(Range("K:K"))
    maxVolume = Application.WorksheetFunction.Max(Range("L:L"))
    
    Set maxPercentRange = ActiveWorkbook.ActiveSheet.Range("K:K").Find(maxPercent, Lookat:=xlWhole)
    Set minPercentRange = ActiveWorkbook.ActiveSheet.Range("K:K").Find(minPercent, Lookat:=xlWhole)
    Set maxVolRange = ActiveWorkbook.ActiveSheet.Range("L:L").Find(maxVolume, Lookat:=xlWhole)
    
   
    'Populate Summary Ticker Values
    Range("P2").Value = maxPercentRange.Offset(, -2)
    Range("P3").Value = minPercentRange.Offset(, -2)
    Range("P4").Value = maxVolRange.Offset(, -3)

End Sub


Sub Format()

    
    'Declare and initialize variables
    Dim myRange1 As Range    'Range to apply formatting
    Dim myRange2 As Range    'Range to apply formatting
    Dim myCondition1, myCondition2, myCondition3, myCondition4 As FormatCondition   'Conditions to meet for formatting
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'Fit all the columns
    Columns("A:T").Select
    Columns("A:T").EntireColumn.AutoFit
    
    'Set the range to format Quantity Change
    Set myRange1 = Range("J2", "J" & lastRow)
    'Set the range to format Percent Change
    Set myRange2 = Range("K2", "K" & lastRow)
    
    
    'Clear any previous formatting in the ranges
     myRange1.FormatConditions.Delete
     myRange2.FormatConditions.Delete
    

    'Set the conditions for Quantity Change - GREEN for positive values and RED for negative values
     Set myCondition1 = myRange1.FormatConditions.Add(xlCellValue, xlGreater, "=0")  'GREEN
     Set myCondition2 = myRange1.FormatConditions.Add(xlCellValue, xlLess, "=0")     'RED
     
     'Positive value - Fill GREEN
     With myCondition1
      .Interior.Color = vbGreen
     End With
     'Negative value - Fill RED
     With myCondition2
      .Interior.Color = vbRed
     End With
     
    'Set the conditions for Percent Change - GREEN for positive values and RED for negative values
     Set myCondition3 = myRange2.FormatConditions.Add(xlCellValue, xlGreater, "=0")  'GREEN
     Set myCondition4 = myRange2.FormatConditions.Add(xlCellValue, xlLess, "=0")     'RED
     
     'Positive value - Fill GREEN
     With myCondition3
      .Interior.Color = vbGreen
     End With
     'Negative value - Fill RED
     With myCondition4
      .Interior.Color = vbRed
     End With
     
    'Hide Calculated Columns
    Columns("R:T").Select
    Selection.EntireColumn.Hidden = True
    
    'Format Percent Change column as Percentage
    Columns("K:K").Select
    Selection.Style = "Percent"
    Range("A1").Select
          
End Sub

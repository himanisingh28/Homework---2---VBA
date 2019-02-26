Attribute VB_Name = "Module3"
Sub multiple_year_stock_date_hard()

Dim stockVolume As Double
Dim i As Long
Dim lRow As Long
Dim ticker As String
Dim rowCounter As Integer
Dim closeValue As Double
Dim openValue As Double
Dim yearlyChange As Double
Dim percentChange As Double

Dim maxPercent As Double
Dim minPercent As Double
Dim greatestTotVol As Double


'Initializing all variables
stockVolume = 0
rowCounter = 2
yearlyChange = 0
percentChange = 0
closeValue = 0
openValue = 0
maxPercent = 0
minPercent = 0
greatestTotVol = 0

'Creating the header values for all columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Value"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'Determining the number of rows in 1st column
lRow = Cells(Rows.Count, 1).End(xlUp).Row


'Executing the loop for all the rows in the 1st column
For i = 2 To lRow
    'To check the first row of a particular ticker value
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value And Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    
    'Store the open value of the ticker in a variable
    openValue = Cells(i, 3).Value
        
    'To check the last row of a particular ticker value
    ElseIf Cells(i - 1, 1).Value = Cells(i, 1).Value And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    stockVolume = Cells(i, 7).Value + stockVolume
    
    'Store the ticker value in a variable
    ticker = Cells(i, 1).Value
    
    'Store the close value of ticker in a variable
    closeValue = Cells(i, 6).Value
        
    'Calculate the yearly change using the open and close value of the ticker
    yearlyChange = closeValue - openValue
    
    'Calculate the percent change and Additional check to handle the divide by 0 error
    If openValue = 0 Then
    percentChange = 0
    Else: percentChange = Round(((yearlyChange / openValue) * 100), 2)
    End If
    
    'Populating the columns with corresponding values
    Cells(rowCounter, 9).Value = ticker
    
    Cells(rowCounter, 10).Value = yearlyChange
    
    Cells(rowCounter, 11).Value = percentChange & "%"
    
    Cells(rowCounter, 12).Value = stockVolume
    
    'Color formatting the yearly change column based on the value
        If Cells(rowCounter, 10).Value >= 0 Then
        Cells(rowCounter, 10).Interior.ColorIndex = 4
    
        ElseIf Cells(rowCounter, 10).Value < 0 Then
        Cells(rowCounter, 10).Interior.ColorIndex = 3
    
        End If
    
    
    rowCounter = rowCounter + 1
    
    stockVolume = 0
    
    Else

    stockVolume = Cells(i, 7).Value + stockVolume
    

    End If
    
    Next i

' To determine stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume"


x = Cells(Rows.Count, 12).End(xlUp).Row


For c = 2 To x

If Cells(c, 11).Value > maxPercent Then

    maxPercent = Cells(c, 11).Value
    Range("P2").Value = Cells(c, 9).Value
    Range("Q2").Value = (maxPercent * 100) & "%"
    
ElseIf Cells(c, 11).Value < minPercent Then

    minPercent = Cells(c, 11).Value
    Range("P3").Value = Cells(c, 9).Value
    Range("Q3").Value = (minPercent * 100) & "%"
    
ElseIf Cells(c, 12).Value > greatestTotVal Then

    greatestTotVal = Cells(c, 12).Value
    Range("P4").Value = Cells(c, 9).Value
    Range("Q4").Value = greatestTotVal

End If
Next c


End Sub


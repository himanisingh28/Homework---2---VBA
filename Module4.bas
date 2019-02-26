Attribute VB_Name = "Module4"
Sub multiple_year_stock_date_challenge()

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

Dim ws As Worksheet

For Each ws In Worksheets

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
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Value"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Determining the number of rows in 1st column
lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Executing the loop for all the rows in the 1st column
For i = 2 To lRow
    'To check the first row of a particular ticker value
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    
    'Store the open value of the ticker in a variable
    openValue = ws.Cells(i, 3).Value
        
    'To check the last row of a particular ticker value
    ElseIf ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    stockVolume = ws.Cells(i, 7).Value + stockVolume
    
    'Store the ticker value in a variable
    ticker = ws.Cells(i, 1).Value
    
    'Store the close value of ticker in a variable
    closeValue = ws.Cells(i, 6).Value
        
    'Calculate the yearly change using the open and close value of the ticker
    yearlyChange = closeValue - openValue
    
    'Calculate the percent change and Additional check to handle the divide by 0 error
    If openValue = 0 Then
    percentChange = 0
    Else: percentChange = Round(((yearlyChange / openValue) * 100), 2)
    End If
    
    'Populating the columns with corresponding values
    ws.Cells(rowCounter, 9).Value = ticker
    
    ws.Cells(rowCounter, 10).Value = yearlyChange
    
    ws.Cells(rowCounter, 11).Value = percentChange & "%"
    
    ws.Cells(rowCounter, 12).Value = stockVolume
    
    'Color formatting the yearly change column based on the value
        If ws.Cells(rowCounter, 10).Value >= 0 Then
        ws.Cells(rowCounter, 10).Interior.ColorIndex = 4
    
        ElseIf ws.Cells(rowCounter, 10).Value < 0 Then
        ws.Cells(rowCounter, 10).Interior.ColorIndex = 3
    
        End If
    
    
    rowCounter = rowCounter + 1
    
    stockVolume = 0
    
    Else

    stockVolume = ws.Cells(i, 7).Value + stockVolume
    

    End If
    
    Next i

' To determine stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume"


x = ws.Cells(Rows.Count, 12).End(xlUp).Row


For c = 2 To x

    If ws.Cells(c, 11).Value > maxPercent Then

        maxPercent = ws.Cells(c, 11).Value
        ws.Range("P2").Value = ws.Cells(c, 9).Value
        ws.Range("Q2").Value = (maxPercent * 100) & "%"
        
    ElseIf ws.Cells(c, 11).Value < minPercent Then

        minPercent = ws.Cells(c, 11).Value
        ws.Range("P3").Value = ws.Cells(c, 9).Value
        ws.Range("Q3").Value = (minPercent * 100) & "%"
        
    End If
Next c

For c = 2 To x
     
    If ws.Cells(c, 12).Value > greatestTotVol Then
        MsgBox ("I am here")
        greatestTotVol = ws.Cells(c, 12).Value
        ws.Range("P4").Value = ws.Cells(c, 9).Value
        ws.Range("Q4").Value = greatestTotVol
        
    End If
Next c

Next ws

End Sub



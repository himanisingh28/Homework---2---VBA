Attribute VB_Name = "Module1"
Sub multiple_year_stock_date_easy()

Dim stockVolume As Double
Dim i As Long
Dim lRow As Long
Dim ticker As String
Dim rowCounter As Integer

'Initialize the variables
stockVolume = 0
rowCounter = 2

'Create the headers for the ticker and total stock volume columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Value"

'Determine the number of rows in 1st column
lRow = Cells(Rows.Count, 1).End(xlUp).Row

'Run the loop for all the rows in the 1st column i.e. ticker column
For i = 2 To lRow
    
    'This condition searches for the last row of a particular ticker value
    If Cells(i - 1, 1).Value = Cells(i, 1).Value And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    stockVolume = Cells(i, 7).Value + stockVolume
    
    ticker = Cells(i, 1).Value
    
    'The corresponding values of ticker and stock volume are written in the respective cells
    Cells(rowCounter, 9).Value = ticker
    
    Cells(rowCounter, 10).Value = stockVolume
    
    rowCounter = rowCounter + 1
    
    stockVolume = 0
    
    Else

    stockVolume = Cells(i, 7).Value + stockVolume
    

    End If
    
Next i


End Sub



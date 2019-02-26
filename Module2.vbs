Attribute VB_Name = "Module2"
Sub multiple_year_stock_date_moderate()

Dim stockVolume As Double
Dim i As Long
Dim lRow As Long
Dim ticker As String
Dim rowCounter As Integer
Dim closeValue As Double
Dim openValue As Double
Dim yearlyChange As Double
Dim percentChange As Double


'Initializing the variables
stockVolume = 0
rowCounter = 2
yearlyChange = 0
percentChange = 0
closeValue = 0
openValue = 0

'Creating headers for the ticker, yearly change, percent change & total stock volume columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Value"


'Determining the number of rows in the 1st column
lRow = Cells(Rows.Count, 1).End(xlUp).Row

'Executing the loop for each row
For i = 2 To lRow
    
    'To check the first row of a particular ticker value'
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value And Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    
    'Store the open value of the ticker in a variable
    openValue = Cells(i, 3).Value
      
    'To check the last row of a particular ticker value'
    ElseIf Cells(i - 1, 1).Value = Cells(i, 1).Value And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    stockVolume = Cells(i, 7).Value + stockVolume
    
    'Store the ticker value in a variable
    ticker = Cells(i, 1).Value
    
    'Store the close value of ticker in a variable
    closeValue = Cells(i, 6).Value
        
    'Calculate the yearly change using the open and close value of the ticker
    yearlyChange = Round((closeValue - openValue), 9)
    
    'Additional check to handle the divide by 0 error
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


End Sub


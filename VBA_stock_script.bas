' ********Must enable mscorlib.dll library to use this VBA ArrayList from Tools bar to execute code successfully*******
Sub runOnWorksheets()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Activate
        ws.Select
        Call stock
    Next
    Application.ScreenUpdating = True
End Sub

Sub stock()
    ' ********Must enable mscorlib.dll library to use this VBA ArrayList from Tools bar to execute code successfully*******
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).row
    Dim tempStore As String
    Dim countTicker As Integer
    
    Dim arrayOpeningPrice As ArrayList
    Set arrayOpeningPrice = New ArrayList
    
    Dim arrayStockVolume As ArrayList
    Set arrayStockVolume = New ArrayList
    
    Dim arrayMaxMin As ArrayList
    Set arrayMaxMin = New ArrayList
    
    Dim arrayTotalVolume As ArrayList
    Set arrayTotalVolume = New ArrayList
    

    row1 = 2
    row2 = 2
    row3 = 2
    row4 = 2
    col1 = 9
    col2 = 10
    col3 = 11
    col4 = 12
    countTicker = 0
    
    totalStockVolume = 0
    
    
    ' Assign column headers to cells
    Cells(1, col1).Value = "Ticker"
    Cells(1, col1 + 1).Value = "Yearly Change"
    Cells(1, col1 + 2).Value = "Percentage Change"
    Cells(1, col1 + 3).Value = "Total Stock Volume"
    

    For I = 2 To LastRow
        tempStore = Cells(I, 1).Value
        ' adding the openingPrice and Stock_Volume cells per ticker to array and  storing the 0th element for yearly change calculation
        arrayOpeningPrice.Add (Cells(I, 3).Value)
        openingPrice = arrayOpeningPrice(0)
        ' adding the Stock_Volume cells per ticker to array
        arrayStockVolume.Add (Cells(I, 7).Value)
        
        ' condition if next ticker in row is equal to previous then go to continue
        If Cells(I + 1, 1).Value = tempStore Then GoTo Continue
              
        Cells(row1, col1).Value = tempStore
        closingPrice = Cells(I, 6).Value
        
        'yearlyChange calculation
        yearlyChange = closingPrice - openingPrice
        Cells(row2, col2).Value = yearlyChange
        
        ' percentageChange calculation
        percentageChange = (closingPrice - openingPrice) / openingPrice * 100 & "%"
        Cells(row3, col3).Value = percentageChange
        
        'for bonus calculation
        percentageChangemaxmin = (closingPrice - openingPrice) / openingPrice * 100
        arrayMaxMin.Add (percentageChangemaxmin)

        'totalStockVolume calculation
        For t = 0 To arrayStockVolume.Count - 1
            totalStockVolume = totalStockVolume + arrayStockVolume.Item(t)
            arrayTotalVolume.Add (totalStockVolume)
        Next
        Cells(row4, col4).Value = totalStockVolume
        totalStockVolume = 0
        
        'for bonus greatedtotalVolume
        
        arrayOpeningPrice.Clear
        arrayStockVolume.Clear
        
        row1 = row1 + 1
        row2 = row2 + 1
        row3 = row3 + 1
        row4 = row4 + 1
                
Continue:
        
    Next I
       
    '''''BONUS'''''
    
    ' Assign column headers to cells
    Cells(1, col1 + 7).Value = "Ticker"
    Cells(1, col1 + 8).Value = "Value"
    
    Cells(2, col1 + 6).Value = "Greatest % Increase"
    Cells(3, col1 + 6).Value = "Greatest % Decrease"
    Cells(4, col1 + 6).Value = "Greatest Total Volume"
    
    ' greatest % Value logic
    
    arrayMaxMin.Sort
    greatest_Decrease = arrayMaxMin.Item(0)
    arrayMaxMin.Reverse
    greatest_Increase = arrayMaxMin.Item(0)
    arrayTotalVolume.Sort
    arrayTotalVolume.Reverse
    arrayTotalVolume (0)
    
    
    Cells(2, 17).Value = Str(greatest_Increase) + "%"
    Cells(3, 17).Value = Str(greatest_Decrease) + "%"
    Cells(4, 17).Value = arrayTotalVolume(0)
      
    ' greatest% Ticker logic
    
    countTicker = (Cells(Rows.Count, "I").End(xlUp).row) - 1
    For x = 2 To countTicker
    
        If Cells(3, 17).Value = Cells(x, 11).Value Then
            greatest_DecreaseTicker = Cells(x, 9).Value
            
        End If
    Next x
            
    For y = 2 To countTicker
    
        If Cells(2, 17).Value = Cells(y, 11).Value Then
            greatest_IncreaseTicker = Cells(y, 9).Value
            
        End If
    Next y
    
    For Z = 2 To countTicker
        If Cells(4, 17).Value = Cells(Z, 12).Value Then
        greatest_TotalVolumeTicker = Cells(Z, 9).Value
        End If
    Next Z
    
    Cells(2, 16).Value = greatest_IncreaseTicker
    Cells(3, 16).Value = greatest_DecreaseTicker
    Cells(4, 16).Value = greatest_TotalVolumeTicker
    Application.ScreenUpdating = True
End Sub















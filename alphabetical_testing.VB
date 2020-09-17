Sub Homework2():
    Dim WS As Worksheet
   
    
    Dim ticker As String
    Dim difference As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim percent_change As String
    Dim stock_totalvol As Long
    Dim Summary_Table_Row As Integer
    Dim percent_increase As String
    Dim percent_decrease As String
    Dim greatest_percent_increase As String
    Dim greatest_percent_decrease As String
    Dim greatest_stock As Long
    
    
    
    
    For Each WS In Worksheets
    WS.Activate
    'Headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Changes"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume x1000"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume x1000"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    Range("J1:L1").EntireColumn.AutoFit
    Range("O2:O4").EntireColumn.AutoFit
    
    Summary_Table_Row = 2
    stock_totalvol = 0
    
    open_price = Cells(2, 3).Value
    
    
    For i = 2 To Cells(1, 1).End(xlDown).Row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            stock_totalvol = (stock_totalvol + (Cells(i, 7).Value / 1000))
            Range("I" & Summary_Table_Row).Value = ticker
            Range("L" & Summary_Table_Row).Value = stock_totalvol
            If open_price <> 0 Then
                difference = (Cells(i, 6).Value - open_price)
                percent_change = FormatPercent((difference / open_price), 2)
                Range("J" & Summary_Table_Row).Value = difference
                Range("K" & Summary_Table_Row).Value = percent_change
                Summary_Table_Row = Summary_Table_Row + 1
                open_price = Cells((i + 1), 3).Value
                stock_totalvol = 0
            Else
                percent_change = "NaN"
            End If
        
        Else
            stock_totalvol = (stock_totalvol + Cells(i, 7).Value) / 1000
        End If

    Next i
    
    For j = 2 To Cells(1, 10).End(xlDown).Row
     If Cells(j, 10).Value >= 0 Then
     Cells(j, 10).Interior.ColorIndex = 4
     Else
     Cells(j, 10).Interior.ColorIndex = 3
     End If
     
     Next j
     

    'Find Stock with Greatest Percent increase
    greatest_percent_increase = FormatPercent(WorksheetFunction.Max(Range(Cells(2, 11), Cells(Cells(1, 11).End(xlDown).Row, 11))), 2)
    'greatest_percent_increase = FormatPercent(WorksheetFunction.Max(Range("K" & Cells(2, 11).End(xlDown).Row)), 2)
    'Find Stock with Greatest Percent decrease
    greatest_percent_decrease = FormatPercent(WorksheetFunction.Min(Range(Cells(2, 11), Cells(Cells(1, 11).End(xlDown).Row, 11))), 2)
    'Find Stock with greatest volume
    greatest_stock = WorksheetFunction.Max(Range(Cells(2, 12), Cells(Cells(1, 12).End(xlDown).Row, 11)))
    
    'Set Values into appropriate cells
    Range("Q2").Value = greatest_percent_increase
    Range("Q3").Value = greatest_percent_decrease
    Range("Q4").Value = greatest_stock
    
    'For loop to find a set ticker for greatest_percent_increase
    For k = 2 To Cells(2, 9).End(xlDown).Row
        If Cells(k, 11).Value = Range("Q2").Value Then
        Range("P2").Value = Cells(k, 9).Value
        Else
        End If
    Next k
    
      'For loop to find a set ticker for greatest_percent_decrease
    For k = 2 To Cells(2, 9).End(xlDown).Row
        If Cells(k, 11).Value = Range("Q3").Value Then
        Range("P3").Value = Cells(k, 9).Value
        Else
        End If
    Next k
    
    'For loop to find a set ticker for stock with greatest volume
    For k = 2 To Cells(2, 9).End(xlDown).Row
        If Cells(k, 12).Value = Range("Q4").Value Then
        Range("P4").Value = Cells(k, 9).Value
        Else
        End If
    Next k
   
        
    'Lol some code I tried to get to work... a beautiful mess
    'Find the row position of the value using match and set the ticker into appropiate cell using cells(row, column).value
        'search for the row position in list of tickers
    'Range("P2").Value = Cells(Application.Match(greatest_percent_increase, Range(Cells(2, 11), Cells(Cells(1, 11).End(xlDown).Row, 11)), 0), 9).Value
    'Range("P3").Value = Cells(greatest_percent_decrease, 11).Value
    'Range("P4").Value = Cells(greatest_stock, 12).Value
        
    
    
    Next WS
    
End Sub



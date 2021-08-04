Attribute VB_Name = "Module1"
Sub multipleworksheets()
Dim xSheet As Worksheet
    Application.ScreenUpdating = False
    For Each xSheet In Worksheets
        xSheet.Select
        Call multiyear_stock
    Next
    Application.ScreenUpdating = True
End Sub
Sub multiyear_stock()

Dim i As Long

tickerrow = 2
tickertype = 0
totalstockvolume = 0
closingprice = 0
yearlychange = 0
percentchange = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Start = 2

For i = 2 To lastrow

' tickertype = Cells(i, 1).Value

totalstockvolume = totalstockvolume + Cells(i, 7).Value

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    Cells(tickerrow, 12) = totalstockvolume
    totalstockvolume = 0
        
        Cells(tickerrow, 9) = Cells(i, 1).Value
        openingprice = Cells(Start, 3).Value
        
        
        closingprice = Cells(i, 6).Value
        yearlychange = closingprice - openingprice
        Cells(tickerrow, 10).Value = yearlychange
        
    Start = i + 1
    
    If openingprice = 0 Then
    
        Cells(tickerrow, 11).Value = 0
    
    Else
        percentchange = (((closingprice / openingprice) * 100) - 100)
        Cells(tickerrow, 11).Value = percentchange
        
    If Cells(tickerrow, 10).Value >= 0 Then
        Cells(tickerrow, 10).Interior.ColorIndex = 4
    
    ElseIf Cells(tickerrow, 10).Value <= 0 Then
        Cells(tickerrow, 10).Interior.ColorIndex = 3
    
    End If
    
    End If
    
    tickerrow = tickerrow + 1
    
    End If
    
    Next i
End Sub

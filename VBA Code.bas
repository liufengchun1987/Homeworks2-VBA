Attribute VB_Name = "Module1"
Sub WorksheetLoop()
    
Dim ws As Worksheet
    
For Each ws In Worksheets
    ws.Activate
    
    Dim i As Long
    Dim column As Integer
    Dim total As Double
    Dim tickercounter As Integer
    Dim LastRow As Long
    Dim openingprice As Double
    Dim closingprice As Double
    Dim yearlychange As Double
    Dim percentchange As String
    
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greates Total Volume"
    
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    column = 1
    tickercounter = 2
    
    openingprice = Cells(2, 3).Value
    For i = 2 To LastRow
        If Cells(i + 1, column).Value = Cells(i, column).Value Then
            total = total + Cells(i, 7).Value
            closingprice = Cells(i + 1, 6).Value
            yearlychange = closingprice - openingprice
            If openingprice <> 0 Then
                percentchange = FormatPercent((closingprice / openingprice) - 1)
            Else
                percentchange = NIL
            End If
                    
        Else
            total = total + Cells(i, 7).Value
            Cells(tickercounter, 9).Value = Cells(i, 1).Value
            Cells(tickercounter, 10).Value = yearlychange
            Cells(tickercounter, 11).Value = percentchange
            Cells(tickercounter, 12).Value = total
            total = 0
            tickercounter = tickercounter + 1
            openingprice = Cells(i + 1, 3).Value
        End If
    Next i
    
    Min = 0
    Max = 0
    Maxvolume = 0
    For i = 2 To LastRow
        If Cells(i, 10) >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        If Cells(i, 11) > Max Then
            Max = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = FormatPercent(Max, 2)
        Else
            If Cells(i, 11) < Min Then
                Min = Cells(i, 11).Value
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = FormatPercent(Min, 2)
            End If
        End If
    
        If Cells(i, 12).Value > Maxvolume Then
            Maxvolume = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = Maxvolume
        End If
    Next i

Next
End Sub
    

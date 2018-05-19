Sub testhw()

Dim ws As Worksheet

For Each ws In Worksheets:
    ws.Activate
    
    Dim stock_name As String
    
    Dim vol_total As Double
    Range("i1").Value = "Ticker name"
    Range("j1").Value = " Total Volume"
    
    vol_total = 0
    
    Dim Summary_Stock_Row As Integer
    Summary_Stock_Row = 2
    
    
    Dim lastrow As Long
    lastrow = Range("A1").End(xlDown).Row
    
    
               For I = 2 To lastrow
       
               If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
       
                   stock_name = Cells(I, 1).Value
       
                   vol_total = vol_total + Cells(I, 7).Value
       
                   Range("I" & Summary_Stock_Row).Value = stock_name
       
                   Range("J" & Summary_Stock_Row).Value = vol_total
       
                   Summary_Stock_Row = Summary_Stock_Row + 1
       
                   vol_total = 0
                   
            Else
       
                   vol_total = volume_total + Cells(I, 7).Value
       
               End If
    
           Next I
    
    Next ws
End Sub

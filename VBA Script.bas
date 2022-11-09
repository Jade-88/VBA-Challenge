Attribute VB_Name = "Module1"
Sub iStock_Data()

    Dim TickerIndex As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Start As Long
    Dim Total As Double
    Dim Change As Double
    Dim YearlyChange As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    ws.Activate
    
    
    
    Start = 2
    OpenPrice = Cells(Start, 3)
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    TickerIndex = 2
    Total = 0
    Change = 0
    YearlyChange = 0
    
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To RowCount
        OpenPrice = Cells(Start, 3)
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Total = Total + Cells(i, 7).Value
            Cells(TickerIndex, 12).Value = Total
            
            Cells(TickerIndex, 9).Value = Cells(i, 1).Value
        
            ClosePrice = Cells(i, 6).Value
            
            
            Change = ClosePrice - OpenPrice
            
            Cells(TickerIndex, 10).Value = Change
            Cells(TickerIndex, 10).NumberFormat = "0.00"
            
            YearlyChange = Change / OpenPrice
            
            ws.Cells(TickerIndex, 11).Value = YearlyChange
            ws.Cells(TickerIndex, 11).NumberFormat = "0.00%"
            
            'Select case state fot the color
            
             '.Interior.ColorIndex =0
             
             Select Case Change
                    Case Is > 0
                    Cells(TickerIndex, 10).Interior.ColorIndex = 4
                    Case Is < 0
                    Cells(TickerIndex, 10).Interior.ColorIndex = 3
                    Case Else
                    Cells(TickerIndex, 10).Interior.ColorIndex = 0
                End Select
    
            TickerIndex = TickerIndex + 1
            
                                
                                
                                
            Total = 0
        
            Start = i + 1
        End If
    
    
        Total = Total + Cells(i, 7).Value
    Next i
    
    Next ws


End Sub


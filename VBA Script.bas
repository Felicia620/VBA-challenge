Attribute VB_Name = "Module1"
Sub VBAChallenge():
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim ticker_symbol As String
        
        
        Dim total_stock_volume As Double
        total_stock_volume = 0
        
        Dim open_price As Double
        open_price = 0
        
        Dim close_price As Double
        close_price = 0
        
        Dim yearly_price_change As Double
        yearly_price_change = 0
        
        Dim percent_price_change As Double
        percent_price_change = 0
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        open_price = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ticker_symbol = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
                yearly_price_change = close_price - open_price
                
                
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                ws.Range("I" & Summary_Table_Row).Value = ticker_symbol
                ws.Range("J" & Summary_Table_Row).Value = yearly_price_change
                
                If yearly_price_change > 0 Then
                
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                ElseIf yearly_price_change < 0 Then
                
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                
                If open_price <> 0 Then
                
                    percent_price_change = (yearly_price_change / open_price) * 100
                    
                Else
                
                End If
                
                ws.Range("K" & Summary_Table_Row).Value = percent_price_change
                ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
                Summary_Table_Row = Summary_Table_Row + 1
                close_price = 0
                yearly_price_change = 0
                percent_price_change = 0
                total_stock_volume = 0
                open_price = ws.Cells(i + 1, 3).Value
                
            Else
            
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
            End If
            
        Next i
    
    Next ws
    
End Sub


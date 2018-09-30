Attribute VB_Name = "Module1"
Sub VBA_Homework_Moderate()
   
    For Each ws In Worksheets
          
    Dim ticker As String
    Dim total_volume As Double
        total_volume = 0
    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim lastrow As Double
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_change"
    ws.Cells(1, 11).Value = "Yearly_percentage"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(1, 9).Font.Bold = True
    ws.Cells(1, 10).Font.Bold = True
    ws.Cells(1, 12).Font.Bold = True
    ws.Cells(1, 11).Font.Bold = True

    Summary_Table_Row = 2
    
    year_open = ws.Cells(2, 3)
         
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                ws.Cells(Summary_Table_Row, 12).Value = total_volume
                    
                year_close = ws.Cells(i, 6).Value
           
                yearly_change = year_close - year_open
                                
                If year_open <> 0 Then
                    year_percent = (year_close - year_open) / year_open
                Else
                    year_percent = 0
                
                End If
                                        
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                ws.Cells(Summary_Table_Row, 10).Value = yearly_change
                ws.Cells(Summary_Table_Row, 11).Value = year_percent
                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                ws.Cells(Summary_Table_Row, 12).Value = total_volume
                
                If yearly_change >= 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                End If
                
                year_open = ws.Cells(i + 1, 3)
                year_close = 0
                year_percent = 0
                Summary_Table_Row = Summary_Table_Row + 1
                total_volume = 0
            
            Else
                total_volume = total_volume + ws.Cells(i, 7).Value

            End If

        Next i
            ws.Columns("A:L").AutoFit
            Summary_Table_Row = 2
Next
End Sub


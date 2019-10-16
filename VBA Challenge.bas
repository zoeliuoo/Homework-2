Attribute VB_Name = "Module3"
Sub solution()
            
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            
            Dim Ticker As String
            Dim Yearly_Change As Double
            Dim Percent_change As Double
            Dim Total_Stock_Volume As Double
            
            Dim summary_table_row As Integer
            summary_table_row = 2
            Total_Stock_Volume = 0
            
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            
                For J = 2 To lastrow

                      If Cells(J + 1, 1).Value <> Cells(J, 1).Value Then
                      
                      stock_last_row = J
                      
                      Ticker = Cells(J, 1).Value
                      Total_Stock_Volume = Total_Stock_Volume + Cells(J, 7).Value
                      
                      Range("I" & summary_table_row).Value = Ticker
                      Range("L" & summary_table_row).Value = Total_Stock_Volume
                      Range("N" & summary_table_row).Value = stock_last_row

                      summary_table_row = summary_table_row + 1
      
                      Total_Stock_Volume = 0

                      Else
                      Total_Stock_Volume = Total_Stock_Volume + Cells(J, 7).Value
                      
                      End If
                      
                Next J
        
End Sub


Sub dif()

        Dim m As Integer
        Cells(1, 14).Value = "1"
        
        lastrow_summary = Cells(Rows.Count, 9).End(3).Row
        For m = 2 To lastrow_summary
        
        CP = Cells(Cells(m, 14).Value, 6).Value
        OP = Cells(Cells(m - 1, 14).Value + 1, 3).Value
        
        Cells(m, 10).Value = CP - OP
        
        If OP <> 0 Then
        Cells(m, 11).Value = (CP - OP) / OP
        
        Else
        End If
        
        Next
        
        Range("K:K").NumberFormatLocal = "0.00%"

        For Each cell In Range("J2:J" & lastrow_summary)
        
        If cell.Value <= 0 Then
            cell.Interior.ColorIndex = 3
            ElseIf cell.Value > 0 Then
            cell.Interior.ColorIndex = 4
        End If
        
        Next
    
        
 
End Sub

Sub max()

    Cells(1, 17).Value = "Ticker"
    Cells(1, 18).Value = "Value"
    Cells(2, 16).Value = "Greatest % increase"
    Cells(3, 16).Value = "Greatest % decrease"
    Cells(4, 16).Value = "Greatest Total Volume"

    GI = Application.WorksheetFunction.max(Range("K:K"))
    Cells(2, 18) = GI
        Range("R2:R3").NumberFormatLocal = "0.00%"
    GD = Application.WorksheetFunction.Min(Range("K:K"))
    Cells(3, 18) = GD
    GV = Application.WorksheetFunction.max(Range("L:L"))
    Cells(4, 18) = GV
    
    Dim i As Integer
    For i = 1 To 3169
        If Cells(2, 18) = Cells(i, 11) Then
           Cells(2, 17).Value = Cells(i, 9).Value
        End If
        
        If Cells(3, 18) = Cells(i, 11) Then
           Cells(3, 17).Value = Cells(i, 9).Value
        End If
        
        If Cells(4, 18) = Cells(i, 12) Then
           Cells(4, 17).Value = Cells(i, 9).Value
        End If
        
    Next

End Sub
    

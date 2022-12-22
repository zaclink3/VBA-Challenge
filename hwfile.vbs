Sub HWAssignment()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim ticker As String
    
    Dim yearlychange As Double
    Dim openprice As Double
    openprice = Cells(2, 3).Value
    
    
    Dim closeprice As Double
    
    
    Dim percentchange As Double
    percentchange = 0
    
    Dim stockvolume As Double
    stockvolume = 0
    
    Dim summary_table_row As Integer
    summary_table_row = 2
    
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ticker = Cells(i, 1).Value
            
            stockvolume = stockvolume + Cells(i, 7).Value
            
            Range("I" & summary_table_row).Value = ticker
            
            Range("L" & summary_table_row).Value = stockvolume
            
            closeprice = Cells(i, 6).Value
            
            yearlychange = (closeprice - openprice)
            
            Range("J" & summary_table_row).Value = yearlychange
            
                If (openprice = 0) Then
                    percentchange = 0
                    
                Else
                percentchange = yearlychange / openprice
                
                End If
                
          Range("K" & summary_table_row).Value = percentchange
          Range("K" & summary_table_row).NumberFormat = "0.00%"
          
          summary_table_row = summary_table_row + 1
          
          stockvolume = 0
          
          openprice = Cells(i + 1, 3)
        
        Else
            
            stockvolume = stockvolume + Cells(i, 7).Value
    
       End If

        Next i
        
        lastrow_summary_table = Cells(Rows.Count, 10).End(xlUp).Row
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
            
         Next i
         
    
    Next ws

End Sub
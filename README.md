# VBA-Challenge
'Start the Challenge
Sub HWAssignment()

'Set Worksheet variable 
    Dim ws As Worksheet
    'set loop for each worksheet and activate it
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
    
    'set the cell values that you are running the code for
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'set more variables
    Dim ticker As String
    Dim yearlychange As Double
    Dim openprice As Double
    'set the open price so you can do your other formulas
    openprice = Cells(2, 3).Value
    
    'set more variables
    Dim closeprice As Double
    
    'set more variables and percent change and stock volume as zero
    Dim percentchange As Double
    percentchange = 0
    
    Dim stockvolume As Double
    stockvolume = 0
    
    'initialize a row for the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    'create a last row formula since the data is so big
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        'start i loop
        For i = 2 To lastrow
        'formula to say if the next ticker is not equal to this ticker and what happens
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'set ticker name
            ticker = Cells(i, 1).Value
            'add stockvolume
            stockvolume = stockvolume + Cells(i, 7).Value
            'input ticker value
            Range("I" & summary_table_row).Value = ticker
            'input stock volume value
            Range("L" & summary_table_row).Value = stockvolume
            'set close price for formulas
            closeprice = Cells(i, 6).Value
            'now that we have set close and open price we can run this formula
            yearlychange = (closeprice - openprice)
            'set yearlychange formula
            Range("J" & summary_table_row).Value = yearlychange
            'if statement just incase open price was 0
                If (openprice = 0) Then
                    percentchange = 0
                    
                Else
                percentchange = yearlychange / openprice
                
                End If
                'percent change defined and setting the number format
          Range("K" & summary_table_row).Value = percentchange
          Range("K" & summary_table_row).NumberFormat = "0.00%"
          'add one to row so it does not overwrite data
          summary_table_row = summary_table_row + 1
          'reset stockvolume
          stockvolume = 0
          'set new open price
          openprice = Cells(i + 1, 3)
        
        Else
            'add to stock volume if ticker symbol is the same
            stockvolume = stockvolume + Cells(i, 7).Value
    
       End If

        Next i
        'another lastrow formula for newly established rows
        lastrow_summary_table = Cells(Rows.Count, 10).End(xlUp).Row
        'another i loop
        For i = 2 To lastrow_summary_table
        'setting the colors specified in the HW
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
            
         Next i
         
    
    Next ws

End Sub


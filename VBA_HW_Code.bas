Attribute VB_Name = "Module1"
Sub stock_data()
'loop thru all sheets

For Each ws In Worksheets

    'activate worksheet
    ws.Activate

    'new columns
    ws.Cells(1, 10).Value = "<ticker list>"
    ws.Cells(1, 11).Value = "<yearly change>"
    ws.Cells(1, 12).Value = "<percent change>"
    ws.Cells(1, 13).Value = "<total stock volume>"
    ws.Cells(1, 17).Value = "TICKER"
    ws.Cells(1, 18).Value = "VALUE"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    
    'declare variable
    Dim ticker_name As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim ticker_total As Double
    ticker_total = 0
    Dim sumtabrow As Integer
    sumtabrow = 2
    Dim i As Long
    Dim lrow As Long
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
   
    'for loop starts here
    For i = 2 To lrow
        
        'if cell DOES NOT equal
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'create ticker list
            ticker_name = Cells(i, 1).Value
        
            'yearly change (please note: i know that my code for calculating open price is incorrect _
            i have tried for over a week to figure it out. i had tutors and bcs sessions, i know that i need to _
            to capture the first day of trading for each ticker, but i cannot figure it out)
            
            open_price = Cells(frow, 3).Value
            close_price = Cells(lrow, 6).Value
            
            yearly_change = close_price - open_price
        
            'percent change
            percent_change = (close_price - open_price) / open_price
        
            'ticker total
            ticker_total = ticker_total + Cells(i, 7).Value
        
            'print ticker list
            Range("J" & sumtabrow).Value = ticker_name
        
            'print yearly change
            Range("K" & sumtabrow).Value = yearly_change
        
            'print percent change
            Range("L" & sumtabrow).Value = percent_change
            Range("L" & sumtabrow).Style = "Percent"
            Range("L" & sumtabrow).NumberFormat = "0.00%"

            'print ticker total
            Range("M" & sumtabrow).Value = ticker_total
            Range("M" & sumtabrow).NumberFormat = "0"
             
            'move to next row in summary table
            sumtabrow = sumtabrow + 1
            
            'reset ticket ticker total
            ticker_total = 0
        
        'if cell DOES equal
        Else
            ticker_total = ticker_total + Cells(i, 7).Value


        End If
        
    Next i
    
    'for loop for conditionals starts here
    For i = 2 To lrow
        
        For j = 11 To 12
        
            If Cells(i, j).Value < 0 Then
        
                Cells(i, j).Interior.ColorIndex = 3
            
            Else
            
                Cells(i, j).Interior.ColorIndex = 4
        
            End If
        
        Next j
        
    Next i
    
'Bonus

'declare variables
Dim max As Double
Dim min As Double
Dim max_volume As Double

 max_percent = WorksheetFunction.max(Range("L:L"))
 min_percent = WorksheetFunction.min(Range("L:L"))
 bigbucks = WorksheetFunction.max(Range("M:M"))
 
    'find max, min, and max volume
    For i = 2 To lrow
        If Cells(i, 12).Value = max_percent Then
            max = Cells(i, 12).Value
            max_name = Cells(i, 10).Value
            
            ElseIf Cells(i, 12).Value = min_percent Then
            min = Cells(i, 12).Value
            min_name = Cells(i, 10).Value
            
            ElseIf Cells(i, 13).Value = bigbucks Then
            max_volume = Cells(i, 13).Value
            maxvol_name = Cells(i, 10).Value
            
            'print max value, min value, max total value
            Cells(2, 17).Value = max_name
            Cells(2, 18).Value = max
            Cells(2, 18).Style = "Percent"
            Cells(2, 18).NumberFormat = "0.00%"
            
            Cells(3, 17).Value = min_name
            Cells(3, 18).Value = min
            Cells(3, 18).Style = "Percent"
            Cells(3, 18).NumberFormat = "0.00%"

            'print min value
            Cells(4, 17).Value = maxvol_name
            Cells(4, 18).Value = max_volume
            Cells(4, 18).NumberFormat = "0"

        End If

    Next i

Next ws

End Sub





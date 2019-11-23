Attribute VB_Name = "Module1"
Sub Ticker()

Dim lastrow As Long
Dim beginning_year_price As Double
Dim end_year_price As Double
Dim stock_tracker As String
Dim j As Double
Dim volume As Double
Dim price_change As Double
Dim p_min As Double
Dim p_max As Double
Dim v_max As Double
Dim p_range As Range
Dim v_range As Range

'----------------------------
'Loop through all sheets
'----------------------------

    For Each ws In Worksheets

        Dim WorksheetName As String
        
        
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        WorksheetName = ws.Name

        beginning_year_price = 0
        end_year_price = 0
        j = 2
        stock_tracker = ws.Cells(2, 1).Value
        beginning_year_price = ws.Cells(2, 3).Value
        end_year_price = 0
        volume = ws.Cells(2, 7).Value

        'label new colums
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"


        For i = 2 To lastrow

               'if different stock
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    end_year_price = ws.Cells(i, 6).Value
                    ws.Cells(j, 9) = stock_tracker
                    ws.Cells(j, 10) = end_year_price - beginning_year_price
                    ws.Cells(j, 12) = volume
                    
                        'if beginning stockprice = 0
                        If beginning_year_price <> 0 Then
                            price_change = (end_year_price - beginning_year_price) / beginning_year_price
                            ws.Cells(j, 11) = Format(price_change, "Percent")
                            
                        Else
                            ws.Cells(j, 11) = "" 'enter null because you can't divide by 0 if beginning_year_price =0
                          
                        End If
            
                    'conditional formatting
                        If price_change > 0 Then 'if price change is positive, color cell green
                            ws.Cells(j, 10).Interior.ColorIndex = 4
                    
                        Else 'if price change is negative, color cell red
                            ws.Cells(j, 10).Interior.ColorIndex = 3
                
                        End If
        
                    'reset next stock values
                    stock_tracker = ws.Cells(i + 1, 1).Value
                    beginning_year_price = ws.Cells(i + 1, 3).Value
                    volume = ws.Cells(i + 1, 7).Value
            
                    'Increment row on summary chart
                    j = j + 1
            
                'if same stock
                Else
                    stock_tracker = ws.Cells(i, 1).Value
                    volume = volume + ws.Cells(i + 1, 7).Value
                    
                End If

        Next i


        'label max and min table
        ws.Range("O1") = "Ticker"
        ws.Range("P1") = "Value"
        ws.Range("N3") = "Greatest % Increase"
        ws.Range("N4") = "Greatest % Decrease"
        ws.Range("N5") = "Greatest Total Volume"

        'get min and max values
        p_max = WorksheetFunction.Max(ws.Range("K:K")) 'get max percent change
        p_max_row = WorksheetFunction.Match(p_max, ws.Range("K:K"), 0) 'get row of max percent change
        ws.Range("O3") = ws.Cells(p_max_row, 9).Value
        ws.Range("P3") = Format(ws.Cells(p_max_row, 11).Value, "Percent")

        p_min = WorksheetFunction.Min(ws.Range("K:K")) 'get min percent change
        p_min_row = WorksheetFunction.Match(p_min, ws.Range("K:K"), 0) 'get row of min percent change
        ws.Range("O4") = ws.Cells(p_min_row, 9).Value
        ws.Range("P4") = Format(ws.Cells(p_min_row, 11).Value, "Percent")

        v_max = WorksheetFunction.Max(ws.Range("L:L")) 'get max volume change
        v_max_row = WorksheetFunction.Match(v_max, ws.Range("L:L"), 0) 'get row of max volume
        ws.Range("O5") = ws.Cells(v_max_row, 9).Value
        ws.Range("P5") = ws.Cells(v_max_row, 12).Value

    Next ws
    MsgBox ("Done.  Have a day :)")


End Sub

Attribute VB_Name = "Module1"
'NOTES: combine the sheet neatly, i.e. move everything together. In class exercise
'we used a counter to move up 1 every if, that counter added 1 to the cell rows
'of the new table


Sub test_mod2()

'Declare some presumptive variables
Dim ticker_name As String
Dim i As Long
Dim lastRow As Long


'Make sure to declare the ws variable
Dim ws As Worksheet

'Summary table row counter
Dim extra_table_row As Integer
extra_table_row = 2


    For Each ws In Worksheets
    
        'Look for last row all sheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Declare a bunch of variables
        'Start with ticker variables
        Dim ticker_open As Double
        Dim ticker_close As Double
        Dim ticker_total As Double
        
        'Make sure ticker_total knows where to start
        ticker_total = 0
                    
        'Now percent and quarterly change
        Dim percent_change As Double
        Dim quart_change As Double
        
        'Print header titles to each sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'Reset the extra_table_row (chatbot suggestion)
        extra_table_row = 2
    
            'From the second row to the last row, For i
            For i = 2 To lastRow
                
                'If the next row is different from the current one
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                  
                    'Save the closing value to a variable
                    ticker_close = ws.Cells(i, 6).Value
                  
                    'Save ticker_name variable to a cell value
                    ticker_name = ws.Cells(i, 1).Value
                      
                    'Add up the final ticker_total (Total Stock Volume)
                    ticker_total = ticker_total + ws.Cells(i, 7).Value
                
                    'Caluculate the Quarterly Change and save to variable
                    quart_change = ticker_close - ticker_open
                    
                    'Calculate the Percent Change and save to variable
                    'Percent change needs a conditional to keep it from dividing by 0
                    If quart_change <> 0 And ticker_open <> 0 Then
                
                        'Calculate the percent_change
                        percent_change = (quart_change / ticker_open)
                        
                    End If
                    
                    'Print Quarterly Change
                    ws.Cells(extra_table_row, 10).Value = quart_change
                    
                    'Print Percent Change
                    ws.Cells(extra_table_row, 11).Value = percent_change
                    
                    'Print the total
                    ws.Cells(extra_table_row, 12).Value = ticker_total
                    
                    'Print the name to the summary table
                    ws.Cells(extra_table_row, 9).Value = ticker_name
                    
                    'Add formatting to the Quarterley Change (green positive, red negative)
                    If quart_change > 0 Then
                        ws.Cells(extra_table_row, 10).Interior.ColorIndex = 4
                    
                    ElseIf quart_change < 0 Then
                        ws.Cells(extra_table_row, 10).Interior.ColorIndex = 3
                    
                    
                    End If
                    
                    'We need to move the counter up 1 (extra_table_row)
                    extra_table_row = extra_table_row + 1
                    
                    'And now reset all the counters and sums
                    ticker_open = 0
                    ticker_close = 0
                    ticker_total = 0
                    ticker_counter = 0
                    quart_change = 0
                    percent_change = 0
                
                Else
                
                    'Add to the ticker_counter
                    ticker_counter = ticker_counter + 1
                        
                    'If it's the first counter
                    If ticker_counter = 1 Then
                            
                        'Save the row's open value to a variable
                        ticker_open = ws.Cells(i, 3).Value
                    
                    End If
                    
                    'Add up the total volume
                    ticker_total = ticker_total + ws.Cells(i, 7).Value
            
                End If
                       
                       
                       
                
            
            Next i
        
        
        'New last row for extra columns
        lastRow_extra = ws.Cells(Rows.Count, 11).End(xlUp).Row
  
        'Fix formatting in Percent Change column
        Dim percent_format_range As Range
        Set percent_format_range = ws.Range("K2:K" & lastRow_extra)
        percent_format_range.NumberFormat = "0.00%"
  
        'Start summary table stuff with some variable
        Dim max_percent_value As Double
        Dim max_percent_row As Long
        Dim max_percent_ticker As String
        
        'Min variables
        Dim min_percent_value As Double
        Dim min_percent_row As Long
        Dim min_percent_ticker As String
        
        'Max total volume variables
        Dim max_total_value As Double
        Dim max_total_row As Long
        Dim max_total_ticker As String
        
        'Max percent value and location
        max_percent_value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow_extra))
        max_percent_row = ws.Cells(WorksheetFunction.Match(max_percent_value, ws.Columns(11), 0), 11).Row
        max_percent_ticker = ws.Range("I" & max_percent_row).Value
        
        'Min percent value and location
        min_percent_value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow_extra))
        min_percent_row = ws.Cells(WorksheetFunction.Match(min_percent_value, ws.Columns(11), 0), 11).Row
        min_percent_ticker = ws.Range("I" & min_percent_row).Value
        
        'Max total volume
        max_total_value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow_extra))
        max_total_row = ws.Cells(WorksheetFunction.Match(max_total_value, ws.Columns(12), 0), 12).Row
        max_total_ticker = ws.Range("I" & max_total_row).Value
        
        'Print table labels
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Print in summary table
        ws.Range("P2").Value = max_percent_ticker
        ws.Range("P3").Value = min_percent_ticker
        ws.Range("P4").Value = max_total_ticker
        
        'Print in summary table
        ws.Range("Q2").Value = max_percent_value
        ws.Range("Q3").Value = min_percent_value
        ws.Range("Q4").Value = max_total_value
        
        'Fix formatting in Summary Table - Value column (except final entry, which is not a percent but a sum)
        Dim sumvalue_format_range As Range
        Set sumvalue_format_range = ws.Range("Q2:Q3")
        sumvalue_format_range.NumberFormat = "00.00%"

        'This works! Just for some reason won't work on in my code for the assignment
        'It's having trouble knowing where to look, something isn't working there...
        'But it's fine here!
        
        'Quick AutoFit (don't like the it leaves the whole table selected, maybe another way?
        ws.Columns("A:Q").AutoFit
        
        
    Next ws


End Sub













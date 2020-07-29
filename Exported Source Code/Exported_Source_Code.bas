Attribute VB_Name = "Module1"
Sub DataCleanup()

'Declare Values
Dim cur_row As Double, last_row As Integer
Dim cur_col As Integer, last_col As Integer
Dim total_row As Double, total_colum As Integer
Dim tker_num As Double    'tker_num is the total data points of the same ticker
Dim total_table_cells As Double
Dim ws As Worksheet

' Work book contains multiple worksheet
'  Idea is to loop through each of them and process data
'  then move on to the next
For Each ws In Worksheets
    ws.Activate
    
    'Clear any previous data in the output table
    Columns("I:ZZ").Delete Shift:=xlToLeft

    'Count total rows and columns of raw data
    total_row = ActiveSheet.UsedRange.Rows.Count
    total_column = ActiveSheet.UsedRange.Columns.Count
    
    'define starting value
    total_stk_vol = Range("G2").Value
    tker_num = 1

    '+++++++++++++ FIRST SUMMARY TABLE +++++++++++++++
    'begin to loop thru each row & process data
    For cur_row = 2 To total_row
        ' the object is to group all rows with same ticker symbol together
        ' to accomplish this, need to determine if ticker is still the same
        ' between the current row and the next row 
        If Cells(cur_row, 1).Value = Cells(cur_row + 1, 1).Value Then
            'if same -> get row value and add ticker count by 1
            ' add the row value to total stock value
            ticker = Cells(cur_row, 1).Value
            tker_num = tker_num + 1
            total_stk_vol = total_stk_vol + Cells(cur_row + 1, 7).Value
            
        Else
            'Once the ticker no longer match -> output that stock to excel table 
            ' defining output "coordinates"
            Set tbl1 = ActiveSheet.Cells(1, total_column + 2)
          
            'Calculate the yearly change
            'need to plus 1 to move down to first cell below header
            Start_Price = Cells(cur_row - tker_num + 1, 3).Value
            End_Price = Cells(cur_row, 3).Value
            yr_change = End_Price - Start_Price
                If Start_Price = 0 Then
                    perc_change = 0
                Else
                    perc_change = FormatPercent(yr_change / Start_Price, 2)
                End If
            
            'Create Table Header
            With tbl1
                Dim header(0 To 4) As String
                Dim i_header As Integer

                'Create header name for output table                           
                header(0) = "Ticker"
                header(1) = "Count"
                header(2) = "Yearly Change"
                header(3) = "Percent Change"
                header(4) = "Total Stock Volume"
                'looping to parse header name into row
                For i_header = 0 To 4
                    .Offset(0, i_header).Value = header(i_header)
                Next i_header
                
                'Table Value: passing value from left to right 
                total_table_cells = Cells(Rows.Count, total_column + 2).End(xlUp).Row
                'ticker label
                .Offset(total_table_cells, 0).Value = ticker
                'total ticker counts
                .Offset(total_table_cells, 1).Value = tker_num
                'yearly change
                .Offset(total_table_cells, 2).Value = yr_change

                'if statement to change cell color based on value
                If yr_change <= 0 Then 
                    'zero or negative results in red background
                    .Offset(total_table_cells, 2).Interior.ColorIndex = 3
                Else 
                    'otherwise, green background
                    .Offset(total_table_cells, 2).Interior.ColorIndex = 4
                End If

                ' passing in value of percentage change 
                .Offset(total_table_cells, 3).Value = perc_change
                ' format output number to comma seperated thousands
                .Offset(total_table_cells, 4).Value = Format(total_stk_vol, "#,###")
            End With
            
        'reset value after export all the value
        'get prepare for the next loop for next ticker
        ticker = Range("A2").Offset(tker_num + 1, 0)
        tker_num = 1
        yr_change = 0
        perc_change = 0
        total_stk_vol = 0
        End If
    'cur_row is not reset as it will be the baseline for the next ticker
    'if reset, the loop will go back to the previous already analyzed ticker
    Next cur_row
    
    
    '+++++++++++++ SECOND SUMMARY TABLE +++++++++++++++
    ' delare variables
    Dim tbl2_ColumnNumber As Long
    Dim tbl2_ColumnLetter As String
    Dim max_perc_val As Single, min_perc_val As Single


    'count new total columns after table 1 created
    new_total_column = ActiveSheet.UsedRange.Columns.Count
    'locating coordinates to make table 2
    'using output table 1 as the reference point
    perc_change_ColumnNumber = new_total_column - 1
    perc_change_ColumnLetter = Split(Cells(1, perc_change_ColumnNumber).Address, "$")(1)
    tabl1_totalrows = Range(perc_change_ColumnLetter & 1, Range(perc_change_ColumnLetter & 2).End(xlDown)).Rows.Count
    
    'expect to space out 2 columns from table 1
    tbl2_ColumnNumber = new_total_column + 2

    'convert To Column Letter
    tbl2_ColumnLetter = Split(Cells(1, tbl2_ColumnNumber).Address, "$")(1)
    
    ' set table 2 anchor point
    Set tbl2 = Range(tbl2_ColumnLetter & "1")

    ' create the second table row and column headers
    tbl2.Offset(0, 1) = "Ticker"
    tbl2.Offset(0, 2) = "Value"
    tbl2.Offset(1, 0) = "Greatest % Increase"
    tbl2.Offset(2, 0) = "Greatest % Decrease"
    tbl2.Offset(3, 0) = "Greatest Total Volume"


    ' find the greatest % change value
    Dim i_find_perc_max As Integer

    ' in order to have compare between different stock
    ' for the first stock -> need a reference point
    max_perc_val = -1E+16
    min_perc_val = 1E+16
    max_stock_vol = -1E+16

    'looping to find max and min stock 
    For i_find_perc = 2 To tabl1_totalrows
    
        ' Find max % change - 1 col to the left of last col on table 1
        If Cells(i_find_perc, new_total_column - 1) > max_perc_val Then
            max_perc_val = Cells(i_find_perc, new_total_column - 1)
            max_perc_ticker = Cells(i_find_perc, new_total_column - 4)
            tbl2.Offset(1, 1) = max_perc_ticker
            tbl2.Offset(1, 2) = FormatPercent(max_perc_val, 2)
        End If
        
        'Find min % change
        If Cells(i_find_perc, new_total_column - 1) < min_perc_val Then
            min_perc_val = Cells(i_find_perc, new_total_column - 1)
            min_perc_ticker = Cells(i_find_perc, new_total_column - 4)
            tbl2.Offset(2, 1) = min_perc_ticker
            tbl2.Offset(2, 2) = FormatPercent(min_perc_val, 2)
        End If
        
        'Find max stock volume - last column of table 1
        If Cells(i_find_perc, new_total_column) > max_stock_vol Then
            max_stock_vol = Cells(i_find_perc, new_total_column)
            max_stock_ticker = Cells(i_find_perc, new_total_column - 4)
            tbl2.Offset(3, 1) = max_stock_ticker
            tbl2.Offset(3, 2) = Format(max_stock_vol, "#,###")
        End If

    ' next row - next stock
    Next i_find_perc

    ' quick format to make sure all data is shown without cutting off due to overflow                
    Columns("I:ZZ").EntireColumn.AutoFit

' After all stocks in this worksheet analyzed, move to the next worksheet
Next ws

'notify user after all process is done by a message box
MsgBox "All Spreadsheets have been successfully analyzed!"

End Sub


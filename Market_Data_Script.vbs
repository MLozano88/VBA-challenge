Sub Stock_Market_Data()
    On Error Resume Next
        '-Creating variables. The i and j variables will be assigned to the nested loop.
        '-The p variable will be used to assign values to the summary tables.
        '-The year open, year close and total volume will help calculate the yearly
        ' change, percent change and total volume.
        Dim ws_count As Integer
        Dim i As Integer
        Dim j As Long
        Dim p As Integer
        Dim year_open As Double
        Dim year_close As Double
        Dim total_volume As Double
        p = 2

        ws_count = ActiveWorkbook.Worksheets.Count
        
        'Loop through the worksheets
        For i = 1 To ws_count
            'Label columns in the summary tables per sheet.
            ActiveWorkbook.Worksheets(i).Cells(1, 9).Value = "Ticker"
            ActiveWorkbook.Worksheets(i).Cells(1, 10).Value = "Yearly Change"
            ActiveWorkbook.Worksheets(i).Cells(1, 11).Value = "Percent Change"
            ActiveWorkbook.Worksheets(i).Cells(1, 12).Value = "Total Stock Volume"
            'Setting the total volume number to start at 0.
            total_volume = 0
    
            'Loop through the market data in every row beneath the headers.
            For j = 2 To ActiveWorkbook.Worksheets(i).Cells.SpecialCells(xlCellTypeLastCell).Row
                '-This conditional will execute once the script determines the next row belongs to a new ticker.
                '-The conditional will retreive the year close value, calculate the yearly change (and format the cell),
                ' calculate the percent change (and format the cell), calculate the total volume and set all the values back
                ' at 0 for the next ticker. The row on the summary table will also shift down by 1. 
                If ActiveWorkbook.Worksheets(i).Cells(j, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j + 1, 1).Value Then
                    year_close = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value

                    ActiveWorkbook.Worksheets(i).Cells(p, 10).Value = year_close - year_open
                    If ActiveWorkbook.Worksheets(i).Cells(p, 10).Value > 0 Then
                        With ActiveWorkbook.Worksheets(i).Cells(p, 10).Interior
                            .ColorIndex = 4
                        End With
                    Else
                        With ActiveWorkbook.Worksheets(i).Cells(p, 10).Interior
                            .ColorIndex = 3
                        End With
                    End If

                    ActiveWorkbook.Worksheets(i).Cells(p, 11).Value = (year_close - year_open) / year_open
                    ActiveWorkbook.Worksheets(i).Cells(p, 11).NumberFormat = "0.00%"
                    year_open = 0
                    year_close = 0

                    ActiveWorkbook.Worksheets(i).Cells(p, 9).Value = ActiveWorkbook.Worksheets(i).Cells(j, 1).Value
                    total_vol = total_vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                    ActiveWorkbook.Worksheets(i).Cells(p, 12).Value = total_vol
                    p = p + 1
                    total_vol = 0

                'This conditonal statement will execute once the script determines the prior row does not belong to
                ' the same ticker. It will pull the year open value, and begin to calculate the total volume.
                ElseIf ActiveWorkbook.Worksheets(i).Cells(j - 1, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j, 1).Value Then

                    year_open = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value

                    total_vol = total_vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value

                'This else statement will execute if none of the prior conditions are met to continue to calculate
                'the total volume.
                Else

                    total_vol = total_vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value

                End If

            Next j

            '-Once the nested loop finishes, the first loop will reset the P value at 2 to keep the
            ' summary tables at the top of the worksheets.
            p = 2
            
        Next i
    
    
    End Sub
    
    
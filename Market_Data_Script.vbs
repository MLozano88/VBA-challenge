Sub Stock_Market_Data()
    On Error Resume Next
        'Creating variables
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
            'Label columns in the summary table per sheet
            ActiveWorkbook.Worksheets(i).Cells(1, 9).Value = "Ticker"
            ActiveWorkbook.Worksheets(i).Cells(1, 10).Value = "Yearly Change"
            ActiveWorkbook.Worksheets(i).Cells(1, 11).Value = "Percent Change"
            ActiveWorkbook.Worksheets(i).Cells(1, 12).Value = "Total Stock Volume"
            
            total_volume = 0
    
            'Loop through the market data
            For j = 2 To ActiveWorkbook.Worksheets(i).Cells.SpecialCells(xlCellTypeLastCell).Row

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
                ElseIf ActiveWorkbook.Worksheets(i).Cells(j - 1, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j, 1).Value Then

                    year_open = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value

                    total_vol = total_vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                Else

                    total_vol = total_vol + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value

                End If

            Next j

            p = 2
            
        Next i
    
    
    End Sub
    
    
Attribute VB_Name = "Module5"
Sub Stocktrackker2():



    ' Decelaring all required variabiles
    Dim Ticker_Symbol As String
    Dim Yearlychange As Double
    Dim PercentChange As Double
    Dim tot_stock_volume As Double
    Dim tablerow As Double
    Dim tablerow1 As Double
    Dim percentageString As String
    Dim open_price As Double
    Dim close_price As Double
    Dim lastRow As Long
    Dim I As Double
    Dim K As Long
    Dim ws As Worksheet
    

    
    
  For Each ws In ActiveWorkbook.Worksheets
  
    ws.Activate
    
    ' function for finding the last row which i found through searching
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Creating columns required for the analysis
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    

    tablerow = 2
    For I = 2 To lastRow
          If ws.Cells(I + 1, 1).Value = ws.Cells(I, 1).Value Then

             tot_stock_volume = tot_stock_volume + ws.Cells(I, 7)
             ws.Cells(tablerow, 12) = tot_stock_volume + ws.Cells(I + 1, 7)

          Else
                Ticker_Symbol = ws.Cells(I, 1)
                ws.Cells(tablerow, 9) = Ticker_Symbol
                tablerow = tablerow + 1
                tot_stock_volume = 0
           End If
            Next I


  ' looping through the tickers to find open and close prices and yearly change


   Dim tablerow2

   K = 2
  For I = 2 To lastRow

          If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

                 ' trying to create a logic where it finds the zero value and skips it
                 For tablerow2 = tablerow To lastRow
                    If ws.Cells(tablerow, 3).Value <> 0 Then
                        tablerow = tablerow2
                      Exit For
                        End If
                   Next tablerow2

                If tablerow <= lastRow Then

                Yearlychange = (ws.Cells(I, 6) - ws.Cells(tablerow, 3))

                If ws.Cells(tablerow, 3).Value <> 0 Then
                    PercentChange = Yearlychange / ws.Cells(tablerow, 3)

                Else
                    PercentChange = 0
                End If

                ws.Cells(K, 10).Value = Yearlychange
                ws.Cells(K, 11).Value = FormatPercent(PercentChange)


                If Yearlychange > 0 Then

                    ws.Cells(K, 10).Interior.ColorIndex = 4

                Else
                    ws.Cells(K, 10).Interior.ColorIndex = 3
                    End If


           tablerow = I + 1
           K = K + 1
           Yearlychange = 0

           End If
             End If
                Next I


        Dim lastrowtable As Long

            lastrowtable = ws.Cells(Rows.Count, 1).End(xlUp).Row


            Dim r As Range
            Dim max_percent As Double
            Dim min_percent As Double
            Dim q As Range
            Dim Max_Tot_Volume As Double


         Set r = ws.Range("K2:K" & lastrowtable)
            max_percent = Application.WorksheetFunction.Max(r)
            ws.Range("P2") = FormatPercent(max_percent)

            min_percent = Application.WorksheetFunction.Min(min_percent)
            ws.Range("P3").Value = FormatPercent(c)

            Set q = ws.Range("L2:L" & lastrowtable)
            Max_Tot_Volume = Application.WorksheetFunction.Max(q)
            ws.Range("P4").Value = Max_Tot_Volume




 Next ws
End Sub



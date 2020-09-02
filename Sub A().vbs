Sub A()
For Each ws In Worksheets

' set variables for holding ticker, open price, close price and total stock volume
Dim summary_row As Long
Dim total_stock_volume As Double
Dim LastRow As Long

' set variable to hold Last Row
Dim WorksheetName As String

' Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Grabbed the WorksheetName
WorksheetName = ws.Name

' Create a Summary Row to hold the Row numbers
summary_row = 2

' Initialise the total stock volume
total_stock_volume = 0

' To calculate each stocks total volume and record the Ticker Name and Total Stock Volume in the corresponding cells
For i = 2 To LastRow

    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        ws.Range("L" & summary_row).Value = total_stock_volume

    Else
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        ws.Range("I" & summary_row).Value = ws.Cells(i, 1).Value
        ws.Range("L" & summary_row).Value = total_stock_volume
        ' Prepare a new row for the next stock
        summary_row = summary_row + 1
        ' Reset the total volume for the next stock to be 0
        total_stock_volume = 0

    End If

Next i


' set variables for value at the year start, year end for the stock as well as the ticker row
Dim year_start_value As Double
Dim year_end_value As Double
Dim ticker_count As Long

' initialise the ticker count
ticker_count = 0

summary_row = 2

' To calculate each stocks total volume and record the yearly change and percent change in the corresponding cells
For i = 2 To LastRow

' when the next cell is a different ticker, then calculate the start year and end year value for the current ticker
year_end_value = ws.Cells(i, 6).Value
year_start_value = ws.Cells(i - ticker_count, 3).Value

  ' When the year start value greater than 0, if they are the same tickers then go to the next cell and count the numbers to use for substracting
   If year_start_value > 0 And ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
      ticker_count = ticker_count + 1

         ' When the year start value greater than 0, if they are the different tickers then calculate the yearly chage price and percentage for the current ticker
        ElseIf year_start_value > 0 And ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

          ' Set two columes to hold the value in correspondant to each ticker
          ws.Range("J" & summary_row) = year_end_value - year_start_value

          ws.Range("K" & summary_row) = ws.Cells(summary_row, 10).Value / year_start_value

          ' A new row for next ticker
          summary_row = summary_row + 1

          ' Reset the ticker count to 0 for the next ticker
          ticker_count = 0

         ' When the year start value equals to 0, if the next ticker is a different one, then record the current ticker yearly change to 0 for both price and percentage
         ElseIf year_start_value = 0 And Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
         ws.Range("J" & summary_row) = 0
         ws.Range("K" & summary_row) = 0

         summary_row = summary_row + 1

    End If

Next i

' set the summary row to hold the value in each line
For summary_row = 2 To LastRow

   ' If the value is negative, then print red colour
    If ws.Range("J" & summary_row) < 0 Then
       ws.Range("J" & summary_row).Interior.ColorIndex = 3

    ' If the value is positive, then print green colour
    ElseIf Range("J" & summary_row) > 0 Then
        ws.Range("J" & summary_row).Interior.ColorIndex = 4

    ' otherwise nothing
    Else

    End If

Next summary_row

' to change the number format in percent change column to percentage style
For i = 2 To LastRow
ws.Cells(i, 11).Value = ws.Cells(i, 11).Value * 100
ws.Cells(i, 11).NumberFormat = "0.00\%"
Next i

Next ws

End Sub

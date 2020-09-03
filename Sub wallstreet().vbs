Sub A()


' set variables for holding ticker, open price, close price and total stock volume
Dim summary_row As Long
Dim total_stock_volume As Double
Dim LastRow As Long


' Determine the Last Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row


' Create a Summary Row to hold the Row numbers
summary_row = 2

' Initialise the total stock volume
total_stock_volume = 0

' To calculate each stocks total volume and record the Ticker Name and Total Stock Volume in the corresponding cells
For i = 2 To LastRow

    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        Range("L" & summary_row).Value = total_stock_volume

    Else
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        Range("I" & summary_row).Value = Cells(i, 1).Value
        Range("L" & summary_row).Value = total_stock_volume
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
year_end_value = Cells(i, 6).Value
year_start_value = Cells(i - ticker_count, 3).Value

  ' When the year start value greater than 0, if they are the same tickers then go to the next cell and count the numbers to use for substracting
   If year_start_value > 0 And Cells(i, 1).Value = Cells(i + 1, 1).Value Then
      ticker_count = ticker_count + 1

         ' When the year start value greater than 0, if they are the different tickers then calculate the yearly chage price and percentage for the current ticker
        ElseIf year_start_value > 0 And Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

          ' Set two columes to hold the value in correspondant to each ticker
          Range("J" & summary_row) = year_end_value - year_start_value

          Range("K" & summary_row) = Cells(summary_row, 10).Value / year_start_value

          ' A new row for next ticker
          summary_row = summary_row + 1

          ' Reset the ticker count to 0 for the next ticker
          ticker_count = 0

         ' When the year start value equals to 0, if the next ticker is a different one, then record the current ticker yearly change to 0 for both price and percentage
         ElseIf year_start_value = 0 And Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
         Range("J" & summary_row) = 0
         Range("K" & summary_row) = 0

         summary_row = summary_row + 1

    End If

Next i

' set the summary row to hold the value in each line
For summary_row = 2 To LastRow

   ' If the value is negative, then print red colour
    If Range("J" & summary_row) < 0 Then
       Range("J" & summary_row).Interior.ColorIndex = 3

    ' If the value is positive, then print green colour
    ElseIf Range("J" & summary_row) > 0 Then
        Range("J" & summary_row).Interior.ColorIndex = 4

    ' otherwise nothing
    Else

    End If

Next summary_row

Dim LastRow2 As Long

LastRow2 = Cells(Rows.Count, 11).End(xlUp).Row

' to change the number format in percent change column to percentage style
For summary_row = 2 To LastRow2
Cells(summary_row, 11).Value = Cells(summary_row, 11).Value * 100
Cells(summary_row, 11).NumberFormat = "0.00\%"
Next summary_row


' --------------------
' CHALLENGE
' --------------------

Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_total_volume As Double
Dim myrange As Range
Dim myrange2 As Range
Dim myrange3 As Range


summary_row = 2

Set myrange = Range("J:J")
Set myrange3 = Range("I:I")

greatest_decrease = Application.WorksheetFunction.Min(myrange)
greatest_increase = Application.WorksheetFunction.Max(myrange)

For summary_row = 2 To LastRow2

If Range("J" & summary_row).Value = greatest_decrease Then
Cells(2, 16).Value = greatest_decrease
Cells(2, 15).Value = Range("I" & summary_row).Value
End If

If Range("J" & summary_row).Value = greatest_increase Then
Cells(3, 16).Value = greatest_increase
Cells(3, 15).Value = Range("I" & summary_row).Value
End If

Next summary_row


Set myrange2 = Range("L:L")

greatest_total_volume = Application.WorksheetFunction.Max(myrange2)

For summary_row = 2 To LastRow2

If Range("L" & summary_row).Value = greatest_total_volume Then
Cells(4, 16).Value = greatest_total_volume
Cells(4, 15).Value = Range("I" & summary_row).Value
End If

Next summary_row


End Sub


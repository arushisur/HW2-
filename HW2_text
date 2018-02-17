Sub wall()
'to loop through all worksheets
Dim Current As Worksheet
For Each Current In Worksheets

'creating headers for all my new columns
Current.Cells(1, 9).Value = "Ticker"
Current.Cells(1, 10).Value = "Yearly Change"
Current.Cells(1, 11).Value = "Percent Change"
Current.Cells(1, 12).Value = "Total Stock Value"
'defining type of variable
Dim vol_row As Integer
Dim last_Row As Long
Dim vol_toal As Double
Dim Closed As Double
Dim Opened As Double
'since row 1 is a header
vol_row = 2
'calculating last row for loop
last_Row = Current.Cells(Rows.Count, 1).End(xlUp).Row

For Row = 2 To last_Row
   If Current.Cells(Row, 1) <> Current.Cells(Row + 1, 1) Then
       Current.Cells(vol_row, 9) = Current.Cells(Row, 1) 'name of ticker
       Current.Cells(vol_row, 12) = vol_total + Current.Cells(Row, 7) 'volume of last similar ticker
       Current.Cells(vol_row, 14) = Current.Cells(Row, 6).Value 'closing value of last similar ticker
       vol_row = vol_row + 1 'need to incriment to next row
       vol_total = 0 ' needs to be set to 0 b/c then it will start adding all the last values of a ticker together
       
   Else
       vol_total = vol_total + Current.Cells(Row, 7)
       
   End If

   If Current.Cells(Row, 1) = Current.Cells(Row + 1, 1) And Current.Cells(Row, 1) <> Current.Cells(Row - 1, 1) Then
   Current.Cells(vol_row, 15) = Current.Cells(Row, 3) 'to find the first opeing value
   
   End If

Next
'I basically created 2 new col that give the opening and closing values for each ticker
'so i am now calculating the num of rows the individual opening or closing values have
'and then looping through that to find the difference and the %, there is also a condition for zeros
Dim Last As Long
Last = Current.Cells(Rows.Count, 15).End(xlUp).Row

For I = 2 To Last
Current.Cells(I, 10).Value = Current.Cells(I, 15) - Current.Cells(I, 14)
Current.Cells(I, 10).NumberFormat = "0.00"

If Current.Cells(I, 10).Value = 0 Then
Current.Cells(I, 11).Value = 0
Else
Current.Cells(I, 11).Value = (Current.Cells(I, 10).Value / Current.Cells(I, 14).Value)
Current.Cells(I, 11).NumberFormat = "0.00%"



End If

If Current.Cells(I, 10).Value > 0 Then
Current.Cells(I, 10).Interior.Color = RGB(0, 255, 0)
Else
Current.Cells(I, 10).Interior.Color = RGB(255, 0, 0)
End If

Next
'message box to check if looping
MsgBox Current.Name
Next

End Sub

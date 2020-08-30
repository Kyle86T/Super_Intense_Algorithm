Sub testing_stock_script()

For Each ws in Worksheets

Dim Ending_Value As Double
Dim Beginning_Value As Double
Dim Count As Double
Dim YoY_Change As Double
Dim YoY_Percent_Change As Double
Dim Ticker As String
Dim Summary_Table_row As Integer
Dim Total_Stock_Volume As Double
Dim LastRow As Long


Total_Stock_Volume = 0
Summary_Table_row = 2
Count = 0
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
'Range("H1") = "Count"
Range("I1") = "Ticker"
Range("J1") = "YoY Change"
Range("K1") = "YoY Percent Change"
Range("L1") = "Total Stock Volume"
'Range("M1") = "Beginning Value"
'Range("N1") = "Ending Value"

For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
Ending_Value = Cells(i, 6).Value
Beginning_Value = Cells(i - Count, 3).Value
YoY_Change = Ending_Value - Beginning_Value
YoY_Percent_Change = (Ending_Value - Beginning_Value) / Beginning_Value
'Range("H" & Summary_Table_row).Value = Count
Range("I" & Summary_Table_row).Value = Ticker
Range("J" & Summary_Table_row).Value = YoY_Change
        If Range("J" & Summary_Table_row).Value > 0 Then
        Range("J" & Summary_Table_row).Interior.ColorIndex = 4
        Else: Range("J" & Summary_Table_row).Interior.ColorIndex = 3
        End If
Range("K" & Summary_Table_row).Value = YoY_Percent_Change
        Range("K" & Summary_Table_row).NumberFormat = "0.00%"
Range("L" & Summary_Table_row).Value = Total_Stock_Volume
'Range("M" & Summary_Table_row).Value = Beginning_Value
'Range("N" & Summary_Table_row).Value = Ending_Value
Summary_Table_row = Summary_Table_row + 1

Count = 0
Beginning_Value = 0
Ending_Value = 0
Total_Stock_Volume = 0


Else
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
Count = Count + 1
Range("H" & Summary_Table_row).Value = Count

End If
Next i
Next ws

End Sub

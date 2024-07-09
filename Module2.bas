Attribute VB_Name = "Module2"
Sub Ticker()
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        outputRow = 2
        
        ws.Cells(1, 9).Value = "Ticker"
      
        For i = 2 To 93001 Step 62
        
            ws.Cells(outputRow, 9).Value = ws.Cells(i, 1).Value
            outputRow = outputRow + 1
        Next i
    Next sheetName
End Sub
Sub QuarterlyChange()
sheetNames = Array("Q1", "Q2", "Q3", "Q4")
 For Each sheetName In sheetNames
    Set ws = ThisWorkbook.Sheets(sheetName)
    outputRow = 2
    ws.Cells(1, 10).Value = "quarterlyChange"
    For i = 2 To 92940 Step 62
        opening = ws.Cells(i, 3).Value
        closing = ws.Cells((i + 61), 6).Value
        change_amt = closing - opening
        ws.Cells(outputRow, 10) = change_amt
        If change_amt < 0 Then
            ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0)
        ElseIf change_amt > 0 Then
            ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0)
        End If
        outputRow = outputRow + 1
    Next i
  Next sheetName
  End Sub
Sub PercentChange()
sheetNames = Array("Q1", "Q2", "Q3", "Q4")
 For Each sheetName In sheetNames
    Set ws = ThisWorkbook.Sheets(sheetName)
    outputRow = 2
    ws.Cells(1, 11).Value = "Percent Change"
    For i = 2 To 92940 Step 62
        opening = ws.Cells(i, 3).Value
        closing = ws.Cells((i + 61), 6).Value
    If opening <> 0 Then
        percent_change = 1 * ((opening - closing) / (-1 * opening))
        ws.Cells(outputRow, 11) = percent_change
        ws.Cells(outputRow, 11).NumberFormat = "0.00"
        outputRow = outputRow + 1
        End If
    Next i
  Next sheetName
End Sub
Sub totalstockvalue()

sheetNames = Array("Q1", "Q2", "Q3", "Q4")
For Each sheetName In sheetNames
Set ws = ThisWorkbook.Sheets(sheetName)
outputRow = 2

ws.Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To 92940 Step 62

vol = 0


vol = vol + ws.Cells(i, 7)
ws.Cells(outputRow, 12) = totalstockVolume


       
   
Next i



Next sheetName
End Sub

Sub summery()

 
sheetNames = Array("Q1", "Q2", "Q3", "Q4")
For Each sheetName In sheetNames
Set ws = ThisWorkbook.Sheets(sheetName)
outputRow = 2

ws.Cells(1, 16).Value = "ticker"
ws.Cells(1, 17).Value = "value"
ws.Cells(2, 15).Value = "greatest%increase"
ws.Cells(3, 15).Value = "greatest%decrease"
ws.Cells(4, 15).Value = "greatesttotalvolume"
For i = 2 To 92940 Step 62

 
 
 Range("Q2") = WorksheetFunction.Max(Range("k2:k92940"))
 Range("P2") = WorksheetFunction.Match(Range("i2,i92940"))
 
 Range("Q3") = WorksheetFunction.Min(Range("k2:k92940"))
 Range("P2") = WorksheetFunction.Match(i, Range("i2,i92940"))
Next i
Next sheetName
End Sub


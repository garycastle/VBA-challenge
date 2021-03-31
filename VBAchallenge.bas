Attribute VB_Name = "Module1"

Sub calculator_()

For Each ws In Worksheets

ws.Range("J1").Value = "Ticker _Symbol"
ws.Range("K1").Value = "Yearly_Change"
ws.Range("L1").Value = "Percent_change"
ws.Range("M1").Value = "Total_Volume"

ticker = ""
nextrow = 2
openamount = 0
closeamount = 0
totalamount = 0
lastrow = (ws.Cells(Rows.Count, 1).End(xlUp).Row)

For i = 2 To lastrow

currentcell = ws.Cells(i, 1).Value
nextcell = ws.Cells(i + 1, 1).Value
lastcell = ws.Cells(i - 1, 1).Value

If (currentcell <> lastcell) Then

totalamount = totalamount + ws.Cells(i, 7).Value
openamount = ws.Cells(i, 3).Value
ticker = currentcell

ElseIf (currentcell = nextcell) Then

totalamount = totalamount + ws.Cells(i, 7).Value
ticker = currentcell

Else
totalamount = totalamount + ws.Cells(i, 7).Value
closeamount = ws.Cells(i, 6).Value

yearlychange = closeamount - openamount

If openamount = 0 Then
percentchange = 0

Else
percentchange = (yearlychange / openamount)

End If

ws.Cells(nextrow, 10).Value = ticker
ws.Cells(nextrow, 11).Value = yearlychange
ws.Cells(nextrow, 12).Value = percentchange
ws.Cells(nextrow, 13).Value = totalamount
nextrow = nextrow + 1
totalamount = 0
End If

Next i

Next ws

End Sub

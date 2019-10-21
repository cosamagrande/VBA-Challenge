'Routine that applies main Sub to all worksheets

Sub AllWorksheets()

Dim wksht As Worksheet
Application.ScreenUpdating = False
For Each wksht In Worksheets
wksht.Select
Call VBA_Challenge

Next
Application.ScreenUpdating = True

End Sub


'Main Routine
Sub VBA_Challenge()

'declaring variables'
Dim Ticker As String
Dim YearOpen As Double
Dim YearClose As Double
Dim Annual_change As Double
Dim Percent_change As Double
Dim Tot_Vol As Double
Dim Sum As Integer



'Setting title cells
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 13).Value = "Annual Open Price"
Cells(1, 14).Value = "Annual Closing Price"

'Define counter that allows to summarize output

Sum = 2

'looking for ticker

lastrow = ActiveSheet.UsedRange.Rows.Count

'Setting counters

Tot_Vol = 0
YearOpen = 0
YearClose = 0
Annual_change = 0
Percent_change = 0

'Locking Annual Opening price for Ticker


For i = 1 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
YearOpen = Cells(i + 1, 3).Value
Cells(Sum, 13).Value = YearOpen
YearOpen = 0
Sum = Sum + 1

End If

Next i

're-setting Sum

Sum = 2

'Locating ticker and closing price

For i = 1 To lastrow

If Cells(i + 1, 1).Value <> Cells(i + 2, 1).Value Then

Ticker = Cells(i + 1, 1).Value
YearClose = Cells(i + 1, 6).Value

'Printing values on Summary
Cells(Sum, 9).Value = Ticker
Cells(Sum, 14).Value = YearClose

'calculating Annual and Percent change and printing values

Annual_change = Cells(Sum, 14).Value - Cells(Sum, 13).Value
Percent_change = Cells(Sum, 13).Value / Cells(Sum, 14).Value
Percent_change = 1 - Percent_change
Cells(Sum, 10).Value = Annual_change
Cells(Sum, 11).Value = Percent_change * 100
Cells(Sum, 12).Value = Tot_Vol


YearClose = 0
Annual_change = 0
Percent_change = 0
Tot_Vol = 0

Sum = Sum + 1


Else
Tot_Vol = Tot_Vol + Cells(i + 2, 7).Value

End If

Next i

'adding conditional coloring to percentage change

For i = 2 To lastrow

If Cells(i, 11).Value < 0 Then

Cells(i, 11).Interior.Color = vbRed

ElseIf Cells(i, 11).Value > 0 Then

Cells(i, 11).Interior.Color = vbGreen

Else
Cells(i, 11).Interior.Color = xlNone

End If

Next i

End Sub



Sub CreditCardForAllSheets()
    Dim ws As Worksheet

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Call the CreditCard subroutine for each worksheet
        CreditCard ws
    Next ws
End Sub

Sub CreditCard(ws As Worksheet)
    ' Set variables for main functionality
    Dim Total As LongLong
    Dim SummaryRow As LongLong
    Dim lastrow As LongLong
    Dim open_1 As Double
    Dim close_1 As Double

    ' Initialize variables for main functionality
    Total = 0
    SummaryRow = 2
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Set variables for tracking greatest % increase, % decrease, and total volume
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As LongLong
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String

    ' Initialize variables for tracking greatest % increase, % decrease, and total volume
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    maxIncreaseTicker = ""
    maxDecreaseTicker = ""
    maxVolumeTicker = ""

    ' Loop through rows in the column for main functionality
    For i = 2 To lastrow
        ' Check if it's a new stock
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' Store the opening price for the new stock
            open_1 = ws.Cells(i, 3).Value
            ' Reset Total for the new stock
            Total = 0
        End If

        ' Accumulate the total stock volume
        Total = Total + ws.Cells(i, 7).Value

        ' Check if it's the last row for the current stock
        If (i = lastrow) Or (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            ' Output results for the current stock
            ws.Cells(SummaryRow, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(SummaryRow, 12).Value = Total
            close_1 = ws.Cells(i, 6).Value
            ws.Cells(SummaryRow, 10).Value = close_1 - open_1
            ws.Cells(SummaryRow, 11).Value = (close_1 / open_1 - 1) * 100
            SummaryRow = SummaryRow + 1
        End If
    Next i

    ' Loop through summary rows to find the greatest % increase, % decrease, and total volume
    For i = 2 To SummaryRow - 1
        ' Check for greatest % increase
        If ws.Cells(i, 11).Value > maxIncrease Then
            maxIncrease = ws.Cells(i, 11).Value
            maxIncreaseTicker = ws.Cells(i, 9).Value
        End If

        ' Check for greatest % decrease
        If ws.Cells(i, 11).Value < maxDecrease Then
            maxDecrease = ws.Cells(i, 11).Value
            maxDecreaseTicker = ws.Cells(i, 9).Value
        End If

        ' Check for greatest total volume
        If ws.Cells(i, 12).Value > maxVolume Then
            maxVolume = ws.Cells(i, 12).Value
            maxVolumeTicker = ws.Cells(i, 9).Value
        End If
    Next i

    ' Output the results for the greatest % increase, % decrease, and total volume at the top
    ws.Cells(1, 15).Value = "Greatest % Increase"
    ws.Cells(1, 16).Value = maxIncreaseTicker
    ws.Cells(1, 17).Value = maxIncrease

    ws.Cells(2, 15).Value = "Greatest % Decrease"
    ws.Cells(2, 16).Value = maxDecreaseTicker
    ws.Cells(2, 17).Value = maxDecrease

    ws.Cells(3, 15).Value = "Greatest Total Volume"
    ws.Cells(3, 16).Value = maxVolumeTicker
    ws.Cells(3, 17).Value = maxVolume
End Sub


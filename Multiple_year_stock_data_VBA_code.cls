VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock()

    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim yearlyChange As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim percentChange As Double
    Dim TotalStockVolume As Double
    Dim summaryRow As Long

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets

        ' Finding the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' start of the summary table row
        summaryRow = 2

        ' creating the headers for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Looping through the worksheet
        For i = 2 To lastRow

            ' finding all data for a ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                ' input the ticker symbol
                ticker = ws.Cells(i, 1).Value

                ' inpupt the opening price for the current ticker
                openingPrice = ws.Cells(i, 3).Value

            End If
            ' calculating the total stock volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            ' checking is its the same ticker or a new ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then

                ' input the closing price for pevious ticker
                closingPrice = ws.Cells(i, 6).Value

                ' Calculating the yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If

                ' Output the ticker and yearly change onto the summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange

                'filling in the yearlychange if its positive or negative
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = vbGreen

                ElseIf yearlyChange <= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = vbRed
                End If

                ' Output the percent change and TotalStockVolume onto the summary table
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = TotalStockVolume

                ' Format the percent change as a percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"

                ' Reset the variables for the next ticker
                TotalStockVolume = 0
                summaryRow = summaryRow + 1

            End If
        Next i

        ' Auto-fit the columns in the summary table
        ws.Columns("I:L").AutoFit

    Next ws
End Sub


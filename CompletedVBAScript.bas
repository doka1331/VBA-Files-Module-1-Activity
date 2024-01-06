
Sub CalculateYearlyChange()
For Each ws In Worksheets
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Use this code to hardcode a variable that represents the last row's number,
'long = VBA integer that is too big for your typical integer/double variable value to hold

    Dim currentTicker As String
    Dim closingPrice As Double
    Dim OpeningPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim TotalStockVolume As Double
    Dim GreatestPercentIncrease As Long
    Dim GreatestPercentDecrease As Long
    Dim GreatestTotalVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim TotalStockTicker As String
    
'Creating variables for use in function; a string = a list of characters displayed as such, nested between quotation marks;
'double =a VBA integer with decimal places
 
    Dim outputRow As Long
    outputRow = 2
    
    Dim OpeningPriceRow As Long
    OpeningPriceRow = 2
'Start writing the output from row 2 (assuming you have headers in row 1)
'Define Where the first opening price is listed in variable OpeningPriceRow, row 2

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = " Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
' Write headers for the output values in columns 9 through 12; and then again for the calculated values output
    GreatestPercentIncrease = 0
    GreatestPercentDecrease = 0
    GreatestTotalVolume = 0
' Loop through the data
    For i = 2 To lastRow
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
'Calculate each ticker symbol's totalstockvolume by adding the previous cell on the conditional that the ticker is the same
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
'if conditional is not met and the ticker symbol in cell (i,1) is not equal to the next cell's ticker symbol/value (aka a new clicker symbol),
'calculate and output the previous data
                 currentTicker = Cells(i, 1).Value
                ws.Cells(outputRow, 9).Value = currentTicker
                closingPrice = ws.Cells(i, 6).Value
                OpeningPrice = ws.Cells(OpeningPriceRow, 3).Value
                yearlyChange = closingPrice - OpeningPrice
                ws.Cells(outputRow, 10).Value = yearlyChange
                ws.Cells(outputRow, 11).Value = Round((yearlyChange / ws.Cells(OpeningPriceRow, 3).Value) * 100, 2)
                If Round((yearlyChange / ws.Cells(OpeningPriceRow, 3).Value) * 100, 2) > GreatestPercentIncrease Then
                    GreatestPercentIncrease = Round((yearlyChange / ws.Cells(OpeningPriceRow, 3).Value) * 100, 2)
                    GreatestIncreaseTicker = ws.Cells(i, 1).Value
                End If
                If Round((yearlyChange / ws.Cells(OpeningPriceRow, 3).Value) * 100, 2) < GreatestPercentDecrease Then
                    GreatestPercentDecrease = Round((yearlyChange / ws.Cells(OpeningPriceRow, 3).Value) * 100, 2)
                    GreatestDecreaseTicker = ws.Cells(i, 1).Value
                End If
                ws.Cells(outputRow, 12).Value = TotalStockVolume
                If TotalStockVolume > GreatestTotalVolume Then
                    GreatestTotalVolume = TotalStockVolume
                    TotalStockTicker = ws.Cells(i, 1).Value
                End If
'Set a conditional format for yearly change that if the value is equal to or greater than 0, cell color = green, otherwise Red
                    If yearlyChange >= 0 Then
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                    End If
'reset totalstockvolume back to Zero for the next clicker value (next i in for loop)
'then add a row to the output row variable
'Find the next row's opening price by adding one to the row section of the formula
                TotalStockVolume = 0
                outputRow = outputRow + 1
                OpeningPriceRow = i + 1
        End If
    Next i
'output calculated values into respective cells
    ws.Range("Q2") = GreatestPercentIncrease
    ws.Range("P2") = GreatestIncreaseTicker
    ws.Range("Q3") = GreatestPercentDecrease
    ws.Range("P3") = GreatestDecreaseTicker
    ws.Range("Q4") = GreatestTotalVolume
    ws.Range("P4") = TotalStockTicker
    
Next ws
End Sub













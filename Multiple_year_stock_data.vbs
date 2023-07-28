' This is the entry point method
Sub Main()
    ' For each sheet in the workbook, call UpdateSummaryForSheet and UpdateTopPerformersForSheet
    For sheetIndex = 1 To Sheets.Count
        UpdateSummaryForSheet sheetIndex
        UpdateTopPerformersForSheet sheetIndex
    Next sheetIndex
    
End Sub

Sub UpdateSummaryForSheet(ByVal sheetIndex As Integer)
    Dim lastDataRow As Long
    
    Dim openPrice As Double
    Dim closePrice As Double
    
    Dim totalVolume As LongLong
    
    Dim prevStockTicker As String
        
    ' Get the last row that has data in it for column A (ticker).  This will be the upper bound of our for loop
    lastDataRow = GetLastRow("A", sheetIndex)
    
    ' If there is stuff in our summary section, clear it out
    ClearSummary lastDataRow, sheetIndex
    
    ' Add the header column text for the summary
    WriteSummaryHeader sheetIndex
    
    prevStockTicker = ""
    
    ' Iterate over all of the records in the data
    For rowCounter = 2 To lastDataRow
    
        ' checks to see if the stock ticker changed since the last time through here
        If (Sheets(sheetIndex).Cells(rowCounter, 1) <> prevStockTicker) Then
            
            ' if this isn't the first time through, we need to write the summary data for the previous ticker
            If (prevStockTicker <> "") Then
                WriteSummaryForTicker prevStockTicker, openPrice, closePrice, totalVolume, sheetIndex
            End If
            
            ' set the new ticker, open and total volume
            prevStockTicker = Sheets(sheetIndex).Cells(rowCounter, 1)
            openPrice = Sheets(sheetIndex).Cells(rowCounter, 3)
            totalVolume = Sheets(sheetIndex).Cells(rowCounter, 7)
        Else
            ' for the same ticker, set the close and add to the total volume
            ' setting the close here, because we don't know if this is the last entry for the ticker until the next time through the loop
            closePrice = Sheets(sheetIndex).Cells(rowCounter, 6)
            totalVolume = totalVolume + Sheets(sheetIndex).Cells(rowCounter, 7)
        End If
        
        ' if this is the last row in the data, we need to write the summary
        If (rowCounter = lastDataRow) Then
            WriteSummaryForTicker prevStockTicker, openPrice, closePrice, totalVolume, sheetIndex
        End If
    Next rowCounter
    
End Sub

Sub UpdateTopPerformersForSheet(ByVal sheetIndex As Integer)
    Dim lastSummaryRow As Long
    
    Dim topIncreaseValue As Double
    Dim topDecreaseValue As Double
    Dim currentPercentChangeValue As Double
    
    Dim topVolumeValue As LongLong
    Dim currentVolumeValue As LongLong
    
    Dim topIncreaseTicker As String
    Dim topDecreaseTicker As String
    Dim topVolumeTicker As String
    Dim currentTicker As String
    
    ' Get the last row of the summary data (column I).  This will be the upper bound of our loop
    lastSummaryRow = GetLastRow("I", sheetIndex)
    
    ' If data is in our top performer area, clean it out
    ClearTopPerformers sheetIndex
    
    ' Write the column headers for the top performers
    WriteTopPerformersHeader sheetIndex
    
    ' initializers
    topIncreaseValue = 0
    topDecreaseValue = 0
    topVolumeValue = 0
    
    topIncreaseTicker = ""
    topDecreaseTicker = ""
    topVolumeTicker = ""
    
    ' For each row in the summary data, keep track of the highest percentage, lowest percentage, and highest volume.
    ' When the criteria is encountered, store the ticker as well
    For rowCounter = 2 To lastSummaryRow
        ' retrieve the current values from the sheet
        currentTicker = Sheets(sheetIndex).Cells(rowCounter, 9)
        currentPercentChangeValue = Sheets(sheetIndex).Cells(rowCounter, 11)
        currentVolumeValue = Sheets(sheetIndex).Cells(rowCounter, 12)
        
        If (currentPercentChangeValue > topIncreaseValue) Then
            topIncreaseTicker = currentTicker
            topIncreaseValue = currentPercentChangeValue
        End If
        
        If (currentPercentChangeValue < topDecreaseValue) Then
            topDecreaseTicker = currentTicker
            topDecreaseValue = currentPercentChangeValue
        End If
        
        If (currentVolumeValue > topVolumeValue) Then
            topVolumeTicker = currentTicker
            topVolumeValue = currentVolumeValue
        End If
    Next rowCounter
    
    ' Report our findings
    Sheets(sheetIndex).Cells(2, 15) = "Greatest % Increase"
    Sheets(sheetIndex).Cells(3, 15) = "Greatest % Decrease"
    Sheets(sheetIndex).Cells(4, 15) = "Greatest Total Volume"
    
    Sheets(sheetIndex).Cells(2, 16) = topIncreaseTicker
    Sheets(sheetIndex).Cells(3, 16) = topDecreaseTicker
    Sheets(sheetIndex).Cells(4, 16) = topVolumeTicker
    
    Sheets(sheetIndex).Cells(2, 17) = FormatPercent(topIncreaseValue, 2)
    Sheets(sheetIndex).Cells(3, 17) = FormatPercent(topDecreaseValue, 2)
    Sheets(sheetIndex).Cells(4, 17) = topVolumeValue
    
End Sub

' This function will find the last row with data in it for a given column and sheet index
Function GetLastRow(ByVal col As String, ByVal sheetIndex As Integer) As Long
    Dim lastRow As Long
    
    lastRow = Sheets(sheetIndex).Range(col & Sheets(sheetIndex).Rows.Count).End(xlUp).Row
    
    GetLastRow = lastRow
End Function

' Setup the column headers for the summary section, given the sheet index
Sub WriteSummaryHeader(ByVal sheetIndex As Integer)
    Sheets(sheetIndex).Cells(1, 9) = "Ticker"
    Sheets(sheetIndex).Cells(1, 10) = "Yearly Change"
    Sheets(sheetIndex).Cells(1, 11) = "Percent Change"
    Sheets(sheetIndex).Cells(1, 12) = "Total Stock Volume"
End Sub

' Write the summary record for the give stock ticker
Sub WriteSummaryForTicker(ByVal stockTicker As String, ByVal openPrice As Double, ByVal closePrice As Double, ByVal totalVolume As LongLong, ByVal sheetIndex As Integer)
    Dim lastSummaryRow As Long
    Dim summaryRow As Long
    
    Dim yearlyChange As Double
    
    ' Finds the last row in the summary section
    lastSummaryRow = GetLastRow("I", sheetIndex)
    
    ' Sets the summary row to the next empty row
    summaryRow = lastSummaryRow + 1
    
    yearlyChange = closePrice - openPrice
    
    Sheets(sheetIndex).Cells(summaryRow, 9) = stockTicker
    Sheets(sheetIndex).Cells(summaryRow, 10) = yearlyChange
    Sheets(sheetIndex).Cells(summaryRow, 11) = FormatPercent(yearlyChange / openPrice, 2)
    Sheets(sheetIndex).Cells(summaryRow, 12) = totalVolume
    
    Sheets(sheetIndex).Cells(summaryRow, 10).Interior.Color = IIf(yearlyChange < 0, vbRed, vbGreen)
    
End Sub

' Setup the column headers for the top performer section, given the sheet index
Sub WriteTopPerformersHeader(ByVal sheetIndex As Integer)
    Sheets(sheetIndex).Cells(1, 16) = "Ticker"
    Sheets(sheetIndex).Cells(1, 17) = "Value"
End Sub

' Clean out the summary section
Sub ClearSummary(ByVal lastDataRow As Long, ByVal sheetIndex As Integer)
    Sheets(sheetIndex).Range("I1:L" & CStr(lastDataRow)).Clear
End Sub

' Clean out the top performer section
Sub ClearTopPerformers(ByVal sheetIndex As Integer)
    Sheets(sheetIndex).Range("O1:Q4").Clear
End Sub




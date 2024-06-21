Attribute VB_Name = "Module1"
Sub YearStockData()

' Set Variables
Dim ws As Worksheet
Dim Ticker_Name As String
Dim lastRow As Long
Dim i As Long
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim QuarterlyChange As Double
Dim PercentChange As Double
Dim Summary_Table_Row As Integer
Dim VolumeT As Double
Dim TableC As Integer
Dim TableD As Integer
Dim TableE As Integer
Dim TableF As Integer
Dim Serie1 As Integer
Dim MinPC As Double
Dim MaxPC As Double
Dim TickerName_MaxDC As String
Dim TickerName_MaxIC As String
Dim MaxVol As Double
Dim TickerName_MaxVol As String

' Establish that the code run in all sheets
For Each ws In ThisWorkbook.Worksheets

' Start Counter
VolumeT = 0
Serie1 = 0
MinPC = 0
MaxPC = 0
MaxVol = 0

TableC = 2
TableD = 2
TableE = 3
TableF = 4

' Set Headers
    Summary_Table_Row = 1
    ws.Range("I" & Summary_Table_Row).Value = "Ticker"
    ws.Range("J" & Summary_Table_Row).Value = "Quarterly Change"
    ws.Range("K" & Summary_Table_Row).Value = "Percent Change"
    ws.Range("L" & Summary_Table_Row).Value = "Total Stock Volume"
    
    ws.Range("P" & Summary_Table_Row).Value = "Ticker"
    ws.Range("Q" & Summary_Table_Row).Value = "Value"
    
    ws.Range("O" & 2).Value = "Greatest % Increase"
    ws.Range("O" & 3).Value = "Greatest % Decrease"
    ws.Range("O" & 4).Value = "Greatest Volume"
    
' Set the Last Row
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Go through the entire row and find the values
    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            If Serie1 = 0 Then
                Serie1 = 1
                OpenPrice = ws.Cells(i, 3).Value
            End If
    
            VolumeT = VolumeT + ws.Cells(i, 7).Value
            
            Else
                ClosePrice = ws.Cells(i, 6).Value
                VolumeT = VolumeT + ws.Cells(i, 7).Value
                QuarterlyChange = ClosePrice - OpenPrice
                ws.Cells(TableC, 10).Value = QuarterlyChange
                PercentChange = QuarterlyChange / OpenPrice
                ws.Cells(TableC, 11).Value = PercentChange
                ws.Cells(TableC, 11).NumberFormat = "0.00%"
                ws.Cells(TableC, 12).Value = VolumeT
                Ticker_Name = ws.Cells(i, 1).Value
                ws.Cells(TableC, 9).Value = Ticker_Name
                
                'Set color for Quarterly Change. Green for positive and red for negative
                If QuarterlyChange < 0 Then
                    ws.Cells(TableC, 10).Interior.ColorIndex = 3
                    ElseIf QuarterlyChange > 0 Then
                        ws.Cells(TableC, 10).Interior.ColorIndex = 4
                End If
                
                TableC = TableC + 1
                
                ' Establish Max Increase, Decrease and Volume
                If PercentChange > MaxPC Then
                    MaxPC = PercentChange
                    TickerName_MaxIC = Ticker_Name
                End If
                If PercentChange < MinPC Then
                    MinPC = PercentChange
                    TickerName_MaxDC = Ticker_Name
                End If
                If VolumeT > MaxVol Then
                    MaxVol = VolumeT
                    TickerName_MaxVol = Ticker_Name
                End If
                
                VolumeT = 0
                Serie1 = 0
                
        End If
    Next i


' Write the results for Max Increase, Decrease and Volume
ws.Cells(TableD, 16).Value = TickerName_MaxIC
ws.Cells(TableD, 17).Value = MaxPC
ws.Cells(TableD, 17).NumberFormat = "0.00%"

ws.Cells(TableE, 16).Value = TickerName_MaxDC
ws.Cells(TableE, 17).Value = MinPC
ws.Cells(TableE, 17).NumberFormat = "0.00%"

ws.Cells(TableF, 16).Value = TickerName_MaxVol
ws.Cells(TableF, 17).Value = MaxVol


Next ws


End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis_all_sheets()

    Dim ws As Worksheet
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Loop through each worksheet (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q#" Then ' Adjust to match the sheet naming convention
            ws.Activate

            ' Set Title Row
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
            
            ' Apply Conditional Formatting for Quarterly Change
            With ws.Range("J2:J" & ws.Rows.Count)
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
                .FormatConditions(1).Interior.Color = RGB(144, 238, 144) ' Light green
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
                .FormatConditions(2).Interior.Color = RGB(255, 182, 193) ' Light red
            End With

            ' Set Initial Values
            j = 0
            total = 0
            greatestIncrease = 0
            greatestDecrease = 0
            greatestVolume = 0
            start = 2

            ' Get the row number of the last row with data
            rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Loop through all rows
            For i = 2 To rowCount
                ' Accumulate total volume
                total = total + ws.Cells(i, 7).Value

                ' If ticker changes, then print results
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ' Calculate change and percent change
                    change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
                    If ws.Cells(start, 3).Value <> 0 Then
                        percentChange = (change / ws.Cells(start, 3).Value) * 100
                    Else
                        percentChange = 0
                    End If

                    ' Print the results
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = change
                    ws.Range("K" & 2 + j).Value = percentChange & "%"
                    ws.Range("L" & 2 + j).Value = total

                    ' Track greatest values
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        greatestIncreaseTicker = ws.Cells(i, 1).Value
                    End If
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        greatestDecreaseTicker = ws.Cells(i, 1).Value
                    End If
                    If total > greatestVolume Then
                        greatestVolume = total
                        greatestVolumeTicker = ws.Cells(i, 1).Value
                    End If

                    ' Reset for next ticker
                    total = 0
                    start = i + 1
                    j = j + 1
                End If
            Next i

            ' Print greatest values
            ws.Range("O2").Value = greatestIncreaseTicker
            ws.Range("P2").Value = greatestIncrease & "%"
            ws.Range("O3").Value = greatestDecreaseTicker
            ws.Range("P3").Value = greatestDecrease & "%"
            ws.Range("O4").Value = greatestVolumeTicker
            ws.Range("P4").Value = greatestVolume

        End If
    Next ws

End Sub


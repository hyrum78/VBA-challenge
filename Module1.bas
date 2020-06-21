Attribute VB_Name = "Module1"
Sub project_loop()
    On Error Resume Next
    
' Define variables
    Dim j As Long
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer
    Dim tot_vol As Double
    Dim year_open As Double
    Dim year_close As Double
    k = 2
    
  
' Create columns I:L headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"


' To sum up range, set vol to 0
        tot_vol = 0
        
' Loop thru sheets, find var & % change of price
        For j = 2 To Cells.SpecialCells(xlCellTypeLastCell).Row
            If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
                year_close = Cells(j, 6).Value
                Cells(k, 10).Value = year_close - year_open
                Cells(k, 11).Value = (year_close - year_open) / year_open
                year_open = 0
                year_close = 0
' Find sum "tot_vol" by symbol
                Cells(k, 9).Value = Cells(j, 1).Value
                tot_vol = tot_vol + Cells(j, 7).Value
                Cells(k, 12).Value = tot_vol
                k = k + 1
                tot_vol = 0
            ElseIf Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
                year_open = Cells(j, 3).Value
' Sum stock volume for each ticker
                tot_vol = tot_vol + Cells(j, 7).Value
            Else
                tot_vol = tot_vol + Cells(j, 7).Value
            End If
        Next j

' Formatting
        For l = 2 To Range("L1").CurrentRegion.Rows.Count
            Cells(l, 11).Style = "Percent"
            If Cells(l, 10).Value > 0 Then
                With Cells(l, 10).Interior
                    .ColorIndex = 4
                End With
            Else
                With Cells(l, 10).Interior
                    .ColorIndex = 3
                    
                End With
            End If
        Next l

End Sub


Attribute VB_Name = "Module2"
Sub CombinedTicker()


Dim i As Double
Dim j As Double
Dim NumberofRows As Double
Dim currentTicker As String
Dim volume As Double
Dim newRow As Long
Dim openingPrice As Double
Dim closingPrice As Double
Dim percentChange As Double

newRow = 2

NumberofRows = Cells(Rows.Count, 1).End(xlUp).Row
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Range("L2:L" & NumberofRows).Style = "Percent"

Cells(1, 13).Value = "Total Stock Volume"

For i = 2 To NumberofRows

    currentTicker = Cells(i, 1).Value
    openingPrice = Cells(i, 3).Value
    volume = Cells(i, 7).Value

    For j = i To NumberofRows
        If (Cells(j + 1, 1).Value = currentTicker) Then
        volume = volume + Cells(j + 1, 7).Value
        Else
            closingPrice = Cells(j, 6).Value
            Cells(newRow, 10).Value = currentTicker
            Cells(newRow, 11).Value = closingPrice - openingPrice
            If (Cells(newRow, 11).Value >= 0) Then
                Cells(newRow, 11).Interior.ColorIndex = 4
            Else
                Cells(newRow, 11).Interior.ColorIndex = 3
            End If
            
            If (openingPrice <> 0) Then
            
                Cells(newRow, 12).Value = ((closingPrice - openingPrice) / openingPrice)
            Else
                Cells(newRow, 12).Value = 0
            End If
            Cells(newRow, 13).Value = volume
            newRow = newRow + 1
            i = j
            Exit For

        End If

    Next j

Next i

LastRowVolume = Cells(Rows.Count, 10).End(xlUp).Row
Range("O2").Value = "Greatest Percent Increase"
Range("O3").Value = "Greatest Percent Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Range("Q2:Q3").Style = "Percent"
Range("Q2").Value = Application.WorksheetFunction.Max(Range("L2:L" & LastRowVolume))
Range("Q3").Value = Application.WorksheetFunction.Min(Range("L2:L" & LastRowVolume))
Range("Q4").Value = Application.WorksheetFunction.Max(Range("M2:M" & LastRowVolume))

Dim temp As Integer
temp = 0


    For k = 2 To LastRowVolume
        If (Cells(k, 12).Value = Range("Q2").Value) Then
            Range("P2").Value = Cells(k, 10).Value
            temp = temp + 1
        ElseIf (Cells(k, 12).Value = Range("Q3").Value) Then
            Range("P3").Value = Cells(k, 10).Value
            temp = temp + 1
        End If
        If temp = 2 Then
            Exit For
        End If
    Next k
    
    For k = 2 To LastRowVolume
        If (Cells(k, 13).Value = Range("Q4").Value) Then
            Range("P4").Value = Cells(k, 10).Value
            Exit For
        End If
    Next k

End Sub


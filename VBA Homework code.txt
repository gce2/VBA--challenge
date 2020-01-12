'Create a script that will loop through all the stocks for one year for each run and take the following information.

Sub VBAWallStreet()

Dim xsheet As Worksheet
For Each xsheet In ThisWorkbook.Worksheets
xsheet.Select

Dim tickerOutput As String
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalStockVolumne As Long
Dim OpeningPreviousTicker As Double
Dim OpeningCurrentTicker As Double
Dim ClosingTicker As Double
Dim i As Long
Dim outputRow As Integer
Dim firstOpen As Double

totalStockVolume = 0
firstOpen = Range("C2").Value

outputRow = 2
inputRow = 2

For inputRow = 2 To Range("A2").End(xlDown).Row
'Add to totalstockvolume
totalStockVolume = totalStockVolume + Cells(inputRow, 7).Value

'Check if ticker is changing
If Cells(inputRow, 1).Value <> Cells(inputRow + 1, 1).Value Then

'Ticker is changing.
    'Copy over the ticker symbol into the tickerOutput column
    Cells(outputRow, 11).Value = Cells(inputRow, 1).Value
    Cells(outputRow, 14).Value = totalStockVolume
    'Yearly change
    Cells(outputRow, 12).Value = Cells(inputRow, 6).Value - firstOpen
    If firstOpen <> 0 Then
        percentChange = (Cells(inputRow, 6).Value - firstOpen) / firstOpen
        Cells(outputRow, 13).Value = percentChange
        If percentChange >= 0 Then
            'Color Cells(outputRow, 13) Green
            Cells(outputRow, 13).Interior.ColorIndex = 4
            
        Else
            'Color Cells(outputRow, 13) Red
             Cells(outputRow, 13).Interior.ColorIndex = 3
        End If
    Else
        Cells(outputRow, 13).Value = "NA"
    End If
        
    firstOpen = Cells(inputRow + 1, 3).Value
    totalStockVolume = 0
    outputRow = outputRow + 1

End If
Next inputRow

    Next xsheet
    
End Sub

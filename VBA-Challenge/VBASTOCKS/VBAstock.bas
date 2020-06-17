Attribute VB_Name = "Module1"
Sub StockData()

'----------------- This to loop thru all sheets

' See this link https://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/

 

'Loop thru all sheets

Dim sheet As Worksheet

Dim starting_sheet As Worksheet

Set starting_sheet = ActiveSheet

 

For Each sheet In ThisWorkbook.Worksheets

    sheet.Activate
'Dim PercentChg As Double
'Dim SummaryTable As Integer
    'SummaryTable = 4
Dim results As Integer
Dim LastRow As Long
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
'Adding Titles to Cells
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "YearlyChange"
Cells(1, 11).Value = "PercentChange"
Cells(1, 12).Value = "TotalStockVolume"

results = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
YearlyChange = 0
PercentChange = 0
TotalStockVolume = 0
'Looping Through one year of Stock


For i = 2 To LastRow
    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Cells(results, 9).Value = Cells(i, 1).Value
    YearlyChange = Cells(i, 6).Value - Cells(i, 3).Value
    Cells(results, 10).Value = YearlyChange
    Cells(results, 12).Value = TotalStockVolume
    
        'Caculate PercentChange
        If (Cells(i, 6).Value = 0 And Cells(i, 3).Value = 0) Then
        PercentChange = 0
        ElseIf (Cells(i, 6).Value <> 0 And Cells(i, 3).Value = 0) Then
        PercentChange = 1
        Else: PercentChange = (Cells(i, 6).Value - Cells(i, 3).Value) / (Cells(i, 3).Value)
        Cells(results, 11).Value = PercentChange
        Cells(results, 11).NumberFormat = "0.00%"
        End If
       
    results = results + 1
    YearlyChange = 0
    PercentChange = 0
    TotalStockVolume = 0
    End If

YearlyChange = YearlyChange + (Cells(i, 10))
If (Cells(i, Column + 10).Value > 0 Or Cells(i, Column + 10).Value = 0) Then
    Cells(i, Column + 10).Interior.ColorIndex = 10
ElseIf Cells(i, Column + 10).Value < 0 Then
    Cells(i, Column + 10).Interior.ColorIndex = 3
End If

Next i

'------------------------This part to go thru all sheets --------------------

sheet.Cells(1, 1) = 1 'This sets cell A1 to each sheet to 1

 

Next

 

starting_sheet.Activate 'Activate the worksheet that was originally active
 

End Sub



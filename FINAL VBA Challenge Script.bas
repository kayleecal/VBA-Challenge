Attribute VB_Name = "Module1"
Sub Module2_Challenge()

For Each ws In Worksheets

Dim i As Variant
Dim WorksheetName As String
    WorksheetName = ws.Name
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Volume As Variant
Dim Company_Row As Long
Dim lastRow As Long
Dim maxPercentage As Double
Dim minPercentage As Double
Dim maxValue As Variant
Dim Open_Value As Double
Dim Close_Value As Double


ws.Cells(1, 10).Value = "Ticker"

ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
ws.Range("J1:M1").Columns.AutoFit
ws.Range("P4").Columns.AutoFit
ws.Range("G:G").Columns.AutoFit

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

Company_Row = 2
Print_Row = 2

Total_Volume = 0


For i = 2 To lastRow

    Open_Value = ws.Cells(Company_Row, 3).Value
    Close_Value = ws.Cells(i, 6).Value
 
 
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        Ticker = ws.Cells(i, 1).Value
        Total_Volume = Total_Volume + (ws.Cells(i, 7).Value)
        Yearly_Change = (Close_Value) - (Open_Value)
        Percent_Change = (Yearly_Change) / (Open_Value)
 
    ws.Range("J" & Print_Row).Value = Ticker
    ws.Range("M" & Print_Row).Value = Total_Volume
    ws.Range("K" & Print_Row).Value = Yearly_Change
    ws.Range("L" & Print_Row).Value = Percent_Change
 
    Company_Row = i + 1
    Print_Row = Print_Row + 1
    Total_Volume = 0
 
    Else
 
         Total_Volume = Total_Volume + ws.Cells(i, 7).Value
 
     End If
 
Next i



For i = 2 To lastRow
    ws.Cells(i, 13).NumberFormat = "General"
Next i
 
For i = 2 To lastRow

ws.Cells(i, 11).NumberFormat = "General"
 
    If ws.Cells(i, 11).Value < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
    Else
        ws.Cells(i, 11).Interior.ColorIndex = 4
    End If
 
Next i
 
For i = 2 To lastRow
 ws.Cells(i, 12).NumberFormat = "0.00%"
Next i
 
For i = 2 To 3
 ws.Cells(i, 18).NumberFormat = "0.00%"
Next i

ws.Cells(4, 18).NumberFormat = "General"


maxPercentage = WorksheetFunction.Max(ws.Range("L:L"))
minPercentage = WorksheetFunction.Min(ws.Range("L:L"))
maxValue = WorksheetFunction.Max(ws.Range("M:M"))
ws.Range("Q1:R4").Columns.AutoFit
ws.Range("R2") = maxPercentage
ws.Range("R3") = minPercentage
ws.Range("R4") = maxValue

For i = 2 To lastRow
    If ws.Cells(i, 12).Value = maxPercentage Then
        ws.Range("Q2") = ws.Cells(i, 10)
    ElseIf ws.Cells(i, 12).Value = minPercentage Then
        ws.Range("Q3") = ws.Cells(i, 10)
    ElseIf ws.Cells(i, 13).Value = maxValue Then
        ws.Range("Q4") = ws.Cells(i, 10)
    End If
 
Next i
 
Next ws

End Sub



Attribute VB_Name = "Module1"
Dim ws As Worksheet

Sub Number_Details_Summary()

For Each ws In Worksheets

Dim WorksheetName As String
WorksheetName = ws.Name
Debug.Print WorksheetName


Dim j As Integer
Dim LR1 As Integer
Dim i As LongLongs
Dim LR As LongLong


Dim Closing_Value As Double
Dim Opening_Value As Double
Dim Yearly_Change As Double

Dim Total_Stock_Volume  As LongLong

'Initial values
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row ' last row in column A
j = 2 'counter
Opening_Value = ws.Cells(2, 3).Value 'first value in open column (column C)
Total_Stock_Volume = 0 ' initial value for the total of stock volume

'Headers
ws.Range("I1").Value = "Ticker"
ws.Range("I1").Font.Bold = True

ws.Range("J1").Value = "Yearly Change"
ws.Range("J1").Font.Bold = True

ws.Range("K1").Value = "Percentage Change"
ws.Range("K1").Font.Bold = True

ws.Range("L1").Value = "Total Stock Volume"
ws.Range("L1").Font.Bold = True


For i = 2 To LR
    
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value 'to get the total of stock volume
    
    
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ws.Range("I" & j).Value = ws.Cells(i, 1).Value 'symbol of ticker
        Closing_Value = ws.Cells(i, 6).Value
        
        
        
        Yearly_Change = Closing_Value - Opening_Value
         
        ws.Cells(j, 10).Value = Yearly_Change
        
        If Yearly_Change < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
        
        End If
         
         Percentage_Change = FormatPercent(Yearly_Change / Opening_Value)
         ws.Cells(j, 11).Value = Percentage_Change
         ws.Cells(j, 12).Value = Total_Stock_Volume
         
         j = j + 1
        
         Opening_Value = ws.Cells(i + 1, 3).Value
         
         Total_Stock_Volume = 0
    
    End If

Next i


LR1 = ws.Cells(Rows.Count, 9).End(xlUp).Row 'last row in the ticker result

'Headers and Titles
ws.Range("P1").Value = "Ticker"
ws.Range("P1").Font.Bold = True

ws.Range("Q1").Value = "Value"
ws.Range("Q1").Font.Bold = True

ws.Range("O2").Value = "Greatest % increase"
ws.Range("O2").Font.Bold = True

ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O3").Font.Bold = True

ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("O4").Font.Bold = True


Greatest_Number = WorksheetFunction.Max(ws.Range("K2:K" & LR))
Smallest_Number = WorksheetFunction.Min(ws.Range("K2:K" & LR))
Greatest_Total_Volume = WorksheetFunction.Max(ws.Range("L2:L" & LR))

ws.Range("Q2").Value = FormatPercent(Greatest_Number)
ws.Range("Q3").Value = FormatPercent(Smallest_Number)
ws.Range("Q4").Value = Greatest_Total_Volume


For j = 2 To LR1

    If ws.Cells(j, 11).Value = ws.Cells(2, 17) Then

    ws.Range("P2").Value = ws.Cells(j, 9).Value

    ElseIf ws.Cells(j, 11).Value = ws.Cells(3, 17) Then

    ws.Range("P3").Value = ws.Cells(j, 9).Value

    ElseIf ws.Cells(j, 12).Value = ws.Cells(4, 17) Then

    ws.Range("P4").Value = ws.Cells(j, 9).Value

    End If

Next j

Worksheets(WorksheetName).Columns("A:Z").AutoFit
Next ws

End Sub




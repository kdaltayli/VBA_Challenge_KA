Attribute VB_Name = "Module1"
Sub VBA_Stock()
Dim a As Double
Dim Total As Double
Dim price_open As Double
Dim price_end As Double
Dim Total_1 As Double

For Each ws In ActiveWorkbook.Worksheets
ws.Activate

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "yearly_change"
Cells(1, 11).Value = "percent_change"
Cells(1, 12).Value = "total_stock_vol"


a = 2
i = 2
Cells(a, 9).Value = Cells(a, 1).Value
price_open = Cells(i, 3).Value

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
    
    If Cells(i, 1).Value = Cells(a, 9).Value Then
    
    Total_1 = Total_1 + Cells(i, 7).Value
    
    price_end = Cells(i, 6).Value

    
    Else
    
        
        Cells(a, 10).Value = price_end - price_open
            If price_open = 0 Then
            
            Cells(a, 11).Value = 0
        Else
        
        Cells(a, 11).Value = (price_end - price_open) / price_open
        
        
        End If
        
        Cells(a, 11).Style = "percent"
        
        If Cells(a, 10).Value >= 0 Then
        
        Cells(a, 10).Interior.Color = vbGreen
         Else
         Cells(a, 10).Interior.Color = vbRed
         
         End If
         
        Cells(a, 12).Value = Total_1
        
    price_open = Cells(i, 3).Value
    
    Total_1 = Cells(i, 7).Value
    
    a = a + 1
    Cells(a, 9).Value = Cells(i, 1).Value
    
    End If
    Next i
    

Columns("I:Q").EntireColumn.AutoFit
Cells(1, 1).Select

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest_%_increase"
Cells(3, 15).Value = "Greatest_%_decrease"
Cells(4, 15).Value = "Greatest_total_volume"

    
Set myrange = Range("K:K")

Greatest_increase = WorksheetFunction.Max(myrange)

Greatest_decrease = WorksheetFunction.Min(myrange)

Cells(2, 17).Value = Greatest_increase

Cells(3, 17).Value = Greatest_decrease

Greatest_ticker = WorksheetFunction.Match(Greatest_increase, myrange, 0)

Cells(2, 16).Value = Cells(Greatest_ticker, 9)

smallest_ticker = WorksheetFunction.Match(Greatest_decrease, myrange, 0)

Cells(3, 16).Value = Cells(smallest_ticker, 9)

Set myrange = Range("L:L")

Greatest_volume = WorksheetFunction.Max(myrange)

Cells(4, 17).Value = Greatest_volume

Greatest_vol_ticker = WorksheetFunction.Match(Greatest_volume, myrange, 0)

Cells(4, 16) = Cells(Greatest_vol_ticker, 9)


Range("Q2,Q3").Select

Range("Q3").Activate

Selection.NumberFormat = "0.00%"

Range("Q4").Select

Range("Q4").Select

Selection.NumberFormat = "0.0000E+00"

    Next ws

End Sub

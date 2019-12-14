Attribute VB_Name = "Module1"
Sub wsloop()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

ws.Activate

Call ticker
Call YearChangeFormat
Call PercFormat
Call GreatestCalc

Next ws

End Sub

Sub ticker()

Dim Ticker_name As String
Dim Open_price As Double
Dim Close_price As Double
Dim Yearly_change As Double
Dim Perc_change As Double
Dim Total_stock_vol As Double
Dim Summary_Table_Row As Integer
Yearly_change = 0
Perc_change = 0
Total_stock_vol = 0

Summary_Table_Row = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Open_price = Cells(2, 3).Value
Cells(1, 10) = "Ticker"
Cells(1, 11) = "Yearly Change"
Cells(1, 12) = "Percent change"
Cells(1, 13) = "Total Stock volume"

For i = 2 To lastrow

   
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Ticker_name = Cells(i, 1).Value
      
      ' Print the Ticker name in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_name
      
      'calculate yearly change
      Close_price = Cells(i, 6)
      Yearly_change = (Close_price - Open_price)
      
      'Print the yearly change value in the Summary Table
      Range("K" & Summary_Table_Row).Value = Yearly_change
      
      If Open_price = 0 Then
      Perc_change = 0
      Else
      
      'Calculate Percent change
      Perc_change = (Yearly_change / Open_price)
      
      End If
      
      'Print the Percent change value in the Summary Table
      Range("L" & Summary_Table_Row).Value = Perc_change
      
      ' Add to the Total stock volume
      Total_stock_vol = Total_stock_vol + Cells(i, 7).Value
      
      'Print the Total stock volume to the Summary Table
      Range("M" & Summary_Table_Row).Value = Total_stock_vol
      
      'Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      'New open price for next ticker
      Open_price = Cells(i + 1, 3).Value
      
      'Reset the Yearly changeTotal stock volume and Percent change
      Yearly_change = 0
      Total_stock_vol = 0
      Perc_change = 0
      
      Else
      
      'Add to the Total stock volume
      Total_stock_vol = Total_stock_vol + Cells(i, 7).Value
      
    End If
    
     
Next i
  
    
End Sub

Sub YearChangeFormat()

Dim Stlastrow As Long

Stlastrow = Cells(Rows.Count, "L").End(xlUp).Row

  For L = 2 To Stlastrow
    If Cells(L, 11).Value > 0 Then
            Cells(L, 11).Interior.ColorIndex = 4
            
        Else
            Cells(L, 11).Interior.ColorIndex = 3
            
        End If
    Next L

    
End Sub

Sub PercFormat()

Dim Stlastrow As Long


Stlastrow = Cells(Rows.Count, "L").End(xlUp).Row

Range("L2:L" & Stlastrow).NumberFormat = "0.00%"
End Sub

Sub GreatestCalc()
Cells(2, 16) = "Greatest % Increase"
Cells(3, 16) = "Greatest % Decrease"
Cells(4, 16) = "Greatest Total Volume"
Cells(1, 17) = "Value"

Dim findmaxInc As Double
Dim findminDec As Double

Stlastrow = Cells(Rows.Count, "K").End(xlUp).Row
findmaxInc = Application.WorksheetFunction.Max(Range("L2:L" & Stlastrow))
findminDec = Application.WorksheetFunction.Min(Range("L2:L" & Stlastrow))
findmaxTvol = Application.WorksheetFunction.Max(Range("M2:M" & Stlastrow))
Cells(2, 17) = findmaxInc
Cells(3, 17) = findminDec
Cells(4, 17) = findmaxTvol

Range("Q2:Q3").NumberFormat = "0.00%"


End Sub

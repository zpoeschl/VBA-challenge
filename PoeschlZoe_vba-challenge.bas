Attribute VB_Name = "Module1"
Sub StockAnalysis()

' Loop through each worksheet
For Each ws In Worksheets

'-------ANALYSIS-----------------------------------------------

' Set variables
Dim LastRow As Double
Dim LastRowSummary As Double
Dim Ticker As String
Dim Open_Price As Double
Open_Price = ws.Cells(2, 3).Value
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Volume As Double
Total_Volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Determine the last data row for analysis purposes
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Determine the last summary table row for formatting purposes
LastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row

' Loop through all ticker symbols
    For i = 2 To LastRow
    
        ' Check if ticker symbol is same and if not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Get ticker
            Ticker = ws.Cells(i, 1).Value
            
            ' Get close price
            Close_Price = ws.Cells(i, 6).Value
            
            ' Calculate yearly price change
            Yearly_Change = (Close_Price - Open_Price)
            
            ' Calculate yearly percentage change
            If Open_Price = 0 Then
            
                Percent_Change = 0
                
            Else
            
                Percent_Change = Yearly_Change / Open_Price
                
            End If
            
            ' Add to stock volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            ' Print ticker symbol to summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            ' Print yearly change to summary table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            ' Print percentage change to summary table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            ' Print the total volume per symbol to the summary table
            ws.Range("L" & Summary_Table_Row).Value = Total_Volume
            
            ' Add one row to the summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset values
            Open_Price = ws.Cells(i + 1, 3)
            Close_Price = 0
            Yearly_Change = 0
            Percent_Change = 0
            Total_Volume = 0
        
        ' If the next row is the same ticker symbol, then:
        Else
                        
            ' Add to the total volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
        End If
    
    Next i
 
'-------FORMATTING---------------------------------------------
        
' Row and column labels
ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

' Column spacing
ws.Cells.EntireColumn.AutoFit

' Conditional color formatting
For i = 2 To LastRowSummary

    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If

Next i

' Cell number formatting
ws.Range("K2:K" & LastRowSummary).NumberFormat = "0.00%"
ws.Range("J2:J" & LastRowSummary).NumberFormat = "0.00"
    
'--------------------------------------------------------------
 
Next ws

End Sub

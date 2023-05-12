Attribute VB_Name = "Module1"
Sub Module2()
'Creating Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Value"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
'Organizing Tickers and Volume Sum

    Dim Ticker As String
    
    Dim Volume_Total As Double
    Volume_Total = 0
    
    Dim Opening_Price As Long
    Opening_Price = 2
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim lastrow As Long
        lastrow = Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            
            YOpen = Cells(Opening_Price, 3).Value
            
            YClose = Cells(i, 6).Value
            
            Yearly_Change = YClose - YOpen
            
            Percentage_Changed = Yearly_Change / YOpen * "100"
            
            Volume_Total = Volume_Total + Cells(i, 7).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker
            
            Range("J" & Summary_Table_Row).Value = Yearly_Change
                
            Range("K" & Summary_Table_Row).Value = Percentage_Changed
            
            Range("L" & Summary_Table_Row).Value = Volume_Total
            
            YOpen = 0
            
            YClose = 0
            
            Yearly_Change = 0
            
            Opening_Price = i + 2
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Volume_Total = 0
            
        
        Else
            Volume_Total = Volume_Total + Cells(i, 7).Value
            
        End If
    
    Next i

End Sub
            


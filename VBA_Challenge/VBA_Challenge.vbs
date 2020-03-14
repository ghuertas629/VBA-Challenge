Attribute VB_Name = "Module1"
Sub StockData()

    For Each ws In Worksheets
    
        Dim Last_Row As Long
        Dim Current_Ticker As String
        Dim Next_Ticker As String
        Dim Ticker_Summary As Integer
        Dim Yearly_Change As Double
        Dim Opening_Value As Double
        Dim Closing_Value As Double
        Dim Total_Volume As Double
        Dim Current_Volume As Double
        Dim Percent_Change As Double
    
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
        Ticker_Summary = 2
        Yearly_Change = 0
        Total_Volume = 0
        Opening_Value = ws.Cells(2, 3).Value
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
    
        For i = 2 To Last_Row

            Current_Ticker = ws.Cells(i, 1).Value
            Next_Ticker = ws.Cells(i + 1, 1).Value
            Closing_Value = ws.Cells(i, 6).Value
            Current_Volume = ws.Cells(i, 7).Value
    
            If Current_Ticker <> Next_Ticker Then
                                
                ws.Cells(Ticker_Summary, 9).Value = Current_Ticker
            
                Total_Volume = Total_Volume + Current_Volume
                
                Yearly_Change = Closing_Value - Opening_Value
                
                ws.Cells(Ticker_Summary, 10).Value = Yearly_Change
                
                ws.Cells(Ticker_Summary, 12).Value = Total_Volume
            
                If Opening_Value = 0 Then
                
                    Percent_Change = Yearly_Change / 1
                    
                Else
                
                    Percent_Change = Yearly_Change / Opening_Value
                    
                End If
                
                ws.Cells(Ticker_Summary, 11).NumberFormat = "0.00%"
                    
                ws.Cells(Ticker_Summary, 11).Value = Percent_Change
            
                Ticker_Summary = Ticker_Summary + 1
                
                Opening_Value = ws.Cells(i + 1, 3).Value
                
                Total_Volume = 0
        
            Else
            
                Yearly_Change = Closing_Value - Opening_Value
            
                Total_Volume = Total_Volume + Current_Volume
            
            End If
    
        Next i
     
        Dim Last_Yearly_Change As Long
    
        Last_Yearly_Change = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
        For i = 2 To Last_Yearly_Change
    
            If ws.Cells(i, 10).Value >= 0 Then
            
                ws.Cells(i, 10).Interior.ColorIndex = 4
        
            Else
            
                ws.Cells(i, 10).Interior.ColorIndex = 3
        
            End If
    
        Next i
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        For i = 2 To Last_Yearly_Change
            
            If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
                
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                
                ws.Range("Q2").NumberFormat = "0.00%"
                
            End If
            
            If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
                
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                
                ws.Range("Q3").NumberFormat = "0.00%"
                
            End If
            
            If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
                
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                
            End If
            
        Next i
        
    Next ws

End Sub

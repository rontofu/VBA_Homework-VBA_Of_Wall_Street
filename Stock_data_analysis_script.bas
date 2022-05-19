Attribute VB_Name = "Module1"
Sub stock_data_analysis()
  For Each ws In Worksheets

    ' Set column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decreace"
        ws.Range("O4").Value = "Greatest Total Volume"

    ' Set initial variables and locations
        Dim Ticker_Name As String
        Dim Stock_Open As Double
        Stock_Open = 0
        Dim Stock_Close As Double
        Stock_Close = 0
        Dim Ticker_Total_Volume As Double
        Ticker_Total_Volume = 0
        Dim Yearly_Stock_Change As Double
        Yearly_Stock_Change = 0
        Dim Yearly_Stock_Change_Percent As Double
        Yearly_Stock_Change_Percent = 0
        Dim Greatest_Increase_Ticker As String
        Dim Greatest_Decrease_Ticker As String
        Dim Greatest_Increase_Ticker_Percent As Double
        Greatest_Increase_Ticker_Percent = 0
        Dim Greatest_Decrease_Ticker_Percent As Double
        Greatest_Decrease_Ticker_Percent = 0
        Dim Greatest_Volume_Ticker As String
        Dim Greatest_Volume As Double
        Greatest_Volume = 0
        Dim Lastrow As Long
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
        
        ' Set stock open value
        Stock_Open = ws.Cells(2, 3).Value
        
        ' Determine last row of worksheet
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through worksheets
        For i = 2 To Lastrow
        
            ' Check if we are still on the same ticker name
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the ticker name starting position
                Ticker_Name = ws.Cells(i, 1).Value
                
                ' Add to the ticker total volume
                Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
                
                ' Calculate yearly stock change and yearly stock change percent
                Stock_Close = ws.Cells(i, 6).Value
                Yearly_Stock_Change = Stock_Close - Stock_Open
                
            If Stock_Open = 0 Then
                Yearly_Stock_Change = 0
                Yearly_Stock_Change_Percent = 0
            Else
                Yearly_Stock_Change_Percent = Yearly_Stock_Change / Stock_Open
                
            End If
                
            ' Assign yearly change percent to column K
            ws.Range("K" & Summary_Table_Row).Value = Yearly_Stock_Change_Percent
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
            ' Assign Ticker name to column I
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
            ' Assign yearly change to column J
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Stock_Change
                
            ' Assign total stock volume to column L
            ws.Range("L" & Summary_Table_Row).Value = Ticker_Total_Volume
                
            ' Use conditional formatting to highlight positive and negative yearly stock change
            If (Yearly_Stock_Change >= 0) Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
                
            ' Add one to the summary table row count and get next stock open value
            Summary_Table_Row = Summary_Table_Row + 1
            Stock_Open = ws.Cells(i + 1, 3).Value
              
            ' Calculate greatest increase/decrease ticker & greatest volume ticker
            If (Yearly_Stock_Change_Percent > Greatest_Increase_Ticker_Percent) Then
                    Greatest_Increase_Ticker_Percent = Yearly_Stock_Change_Percent
                    Greatest_Increase_Ticker = Ticker_Name
                    
            ElseIf (Yearly_Stock_Change_Percent < Greatest_Decrease_Ticker_Percent) Then
                    Greatest_Decrease_Ticker_Percent = Yearly_Stock_Change_Percent
                    Greatest_Decrease_Ticker = Ticker_Name
                    
            End If
                       
            If (Ticker_Total_Volume > Greatest_Volume) Then
                Greatest_Volume = Ticker_Total_Volume
                Greatest_Volume_Ticker = Ticker_Name
            End If
                
            ' Reset values
            Yearly_Stock_Change_Percent = 0
            Ticker_Total_Volume = 0
                
            ' Find next ticker total volume
            
            Else
                
                Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
                
            End If
          
        Next i

                'Assign values into corresponding cells
                ws.Range("P2").Value = Greatest_Increase_Ticker
                ws.Range("Q2").Value = Greatest_Increase_Ticker_Percent
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("P3").Value = Greatest_Decrease_Ticker
                ws.Range("Q3").Value = Greatest_Decrease_Ticker_Percent
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P4").Value = Greatest_Volume_Ticker
                ws.Range("Q4").Value = Greatest_Volume
                
                'Adjust column width
                ws.Columns("I:Q").AutoFit
           
     Next ws
End Sub

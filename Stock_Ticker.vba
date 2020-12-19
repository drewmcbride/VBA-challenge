Sub Stock_Ticker()
    
    ' Loop through all sheets
    For Each ws In Worksheets

        lastRowYear = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
        ' Set an initial variable for holding the ticker name
        Dim Ticker_Name As String
    
        ' Set an initial variable for holding the total volume per ticker
        Dim Ticker_Total As Double
    
        ' Set an initial variable for holding the first stock price of year (to get change)
        Dim Opening_Value As Double
        
        ' Set an initial variable for holding the last stock price of year (to get change)
        Dim Closing_Value As Double
        
        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Dim Percent_Change As Double
        
        ' Label Summary Table Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        Ticker_Total = 0
        Opening_Value = ws.Cells(2, 3).Value
        Closing_Value = 0
        
        ' Loop through all ticker transactions
        For i = 2 To lastRowYear
    
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set Closing_Value
                Closing_Value = ws.Cells(i, 6).Value
    
                ' Set the ticker name
                Ticker_Name = ws.Cells(i, 1).Value
    
                ' Add to the Ticker Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
                ' Print the ticker name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    
                ' Print the Brand Amount to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
                
                ' Print the Yearly Change
                ws.Range("J" & Summary_Table_Row).Value = Closing_Value - Opening_Value
                    
                ' Conditional Formatting
                     If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                
                ' Calculating Percentage Change
                    If Opening_Value = 0 Then
                        ws.Range("K" & Summary_Table_Row).Value = 0
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    Else
                        ws.Range("K" & Summary_Table_Row).Value = (Closing_Value - Opening_Value) / Opening_Value
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    End If
                
                ' Set Opening_Value for *next ticker*
                Opening_Value = ws.Cells(i + 1, 3).Value
                
                Ticker_Total = 0
    
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
    
            ' If the cell immediately following a row is the same ticker name...
            Else
    
                ' Add to the Ticker Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
            End If
        
        Next i
    
    Next ws


End Sub



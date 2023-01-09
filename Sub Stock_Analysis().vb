Sub Stock_Analysis()

'Declare and set worksheet
Dim ws As Worksheet

    'Looping through each sheet (Year)
    For Each ws In Worksheets

        'Formatting Summary Table Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Worksheet Variables
        Dim Lastrow As Long
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim l As Long
        
        'Define ticker
        Dim Ticker As String
        Ticker = " "
        Dim Stock_Count As Integer
        Stock_Count = 0
        
        'Define Yearly Change and Percent Change
        Dim Yearly_Change As Double
        Dim Annual_Open As Double
        Dim Daily_High As Double
        Dim Daily_Low As Double
        Dim Annual_Close As Double
        Dim PercentChange As Variant
    
        'Define Total Stock Volume
        Dim StockVokume_Total As Double
        StockVolume_Total = 0

        'Set an initial variable for holding the stock name
        Dim Stock_Name As String
        
        'Keep track of the location for each stock ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        'Last Row
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set Opening Price - This is the first stock's opening price
        Annual_Open = ws.Cells(2, 3).Value

            'Loop through all Stocks
            For i = 2 To Lastrow
    
                'Check if the stock is still within the same stock, if it is not, do:
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set the Stock name, Ticker Symbol
                Ticker = ws.Cells(i, 1).Value
        
                'Add to the stock volume total
                StockVolume_Total = StockVolume_Total + ws.Cells(i, 7).Value

                'Last row of the current stock. Assign it to the annual close price
                Annual_Close = ws.Cells(i, 6).Value
                
                'Calculate yearly change
                Yearly_Change = Annual_Close - Annual_Open
                
                    'Conditoinal Formatting - Yearly Change
                    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                    Else: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                    End If
               
                'Calculate percent change
                If Annual_Open <> 0 Then
                PercentChange = (Yearly_Change / Annual_Open)
                Else
                PercentChange = "NULL"
                End If
               
                'Format Percent Change as a Percentage
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'Print the Ticker symbol in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
        
                'Print the stock volume total to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = StockVolume_Total
    
                'Print the yearly change to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Print the Percent Chnage to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
    
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
          
                'Reset the stock volume Total
                StockVolume_Total = 0
                
                'i + 1 is the first tow of the next stock so assigen its new value
                Annual_Open = Cells(i + 1, 3).Value
                
                ' If the cell immediately following a row is the same ticker symbol...
                Else
        
                ' Add to the stock Total
                StockVolume_Total = StockVolume_Total + ws.Cells(i, 7).Value

                End If

            Next i

        'Greatest Volume - declare a variable
        Dim Max As WorksheetFunction
        Dim Min As WorksheetFunction
        
        Dim Rng As Range
        
        'Greatest Volume - Ticker and Value
        
        Greatest_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
        Set Rng = ws.Range("L:L").Find(Greatest_Volume, LookAt:=xlWhole)
        ws.Range("Q4") = Greatest_Volume

        For l = 2 To Lastrow
        
            If ws.Cells(l, 12).Value = Greatest_Volume Then
        
            Volume = ws.Cells(l, 12).Offset(, -3)
            ws.Range("P4") = Volume
        
            End If
        
        Next l
        

        'Greatest Increase - Ticker and Value

        Greatest_Increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
        Set Rng = ws.Range("K:K").Find(Greatest_Increase, LookAt:=xlWhole)
        ws.Range("Q2") = Greatest_Increase

        For j = 2 To Lastrow
        
            If ws.Cells(j, 11).Value = Greatest_Increase Then
        
            Increase = ws.Cells(j, 12).Offset(, -3)
            ws.Range("P2") = Increase
        
            End If
        
        Next j
        
        'Greatest Decrease - Ticker and Value
        Greatest_Decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
        Set Rng = ws.Range("K:K").Find(Greatest_Decrease, LookAt:=xlWhole)
        ws.Range("Q3") = Greatest_Decrease
        
        For k = 2 To Lastrow
        
            If ws.Cells(k, 11).Value = Greatest_Decrease Then
        
            Decrease = ws.Cells(k, 12).Offset(, -3)
            ws.Range("P3") = Decrease
        
            End If
        
        Next k

        'Format Greatest Increase/Decrease
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws
End Sub




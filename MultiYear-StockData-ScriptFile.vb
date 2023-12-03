Sub MultiYearStockData()

'Loop through Worksheets
    For Each ws In Worksheets

    'Declare Variables
        Dim Ticker_Name As String
        Dim Ticker_Total As Double
        Dim Ticker_Open As Double
        Dim Ticker_Close As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        
        'Variable for total count of the Total_Stock_Volume
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        Dim Summary_Table As Integer
        'Starting Row for Summary_Table
        Summary_Table = 2
       
        'Establish a base Ticker_Open (opening price). Other opening prices will be calculated in the conditional loop
        Ticker_Open = ws.Cells(2, 3).Value

        'Print Summary_Table Column Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Count # of Rows in First Column
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through all rows by the Ticker_Name

        For i = 2 To lastrow

            'Detects when the value of the next cell is different than that of the current cell. In this case, checks if Ticker_Name is different from the next one.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Define Ticker_Name
              Ticker_Name = ws.Cells(i, 1).Value

              'Add the Total_Stock_Volume for the same Ticker_Name
              Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

              'Print Ticker_Name in Summary_Table
              ws.Range("I" & Summary_Table).Value = Ticker_Name

              'Print Total_Stock_Volume for each Ticker_Name in Summary_Table
              ws.Range("L" & Summary_Table).Value = Total_Stock_Volume

              'Gather Data for Ticker_Close (closing price)
              Ticker_Close = ws.Cells(i, 6).Value

              'Calculate Yearly_Change. Yearly_Change is (Ticker_Close - Ticker_Open) (closing price minus the opening price)
               Yearly_Change = (Ticker_Close - Ticker_Open)
              
              'Print Yearly_Change for each Ticker_Name in Summary_Table
              ws.Range("J" & Summary_Table).Value = Yearly_Change

              'Calculate Percent_Change but first check for the non-divisibilty condition.
              'If the Ticker_Open (opening price) a.k.a. "Old Value" is zero, it would lead to a division by zero error.
              'To avoid this, use a non-divisibility condition to check if the denominator is zero before performing the division.
              
              'To calculate Percent_Change: (Closing Price - Opening Price)/Opening Price)*100
              '(Ticker_Close - Ticker_Open)/Ticker_Open)*100
              
              'Since Yearly_Change = (Ticker_Close - Ticker_Open) you can also simply write, Percent_Change = (Yearly_Change/Ticker_Open)
              
              'No need to multiply (Yearly_Change/Ticker_Open) by 100 in this case
              'because the Range().NumberFormat will print the results as a percentage in column "K"
              
                If Ticker_Open = 0 Then
                   Percent_Change = 0
                
                Else
                   Percent_Change = (Yearly_Change / Ticker_Open)
                
                End If

              'Print Percent_Change for each Ticker_Name in Summary_Table
              ws.Range("K" & Summary_Table).Value = Percent_Change
              ws.Range("K" & Summary_Table).NumberFormat = "0.00%"
   
              'Reset Row Counter. Add 1 to Summary_Table
              Summary_Table = Summary_Table + 1

              'Reset Total_Stock_Volume to Zero
              Total_Stock_Volume = 0

              'Reset Ticker_Open (opening price)
              Ticker_Open = ws.Cells(i + 1, 3)
            
            Else
              
              'Add the Total Stock Volume
              Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    'Find last row of Summary_Table
    Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Yearly_Change Color Code Conditional Formatting
        For i = 2 To Summary_Table
            
            If ws.Cells(i, 10).Value > 0 Then
               ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

    'Print Headers for "Greatest" Summary Table columns and rows

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    'To calculate Greatest % Increase and Greatest % Decrease, first find max and min values in column "Percent Change"
    'To calculate Greatest Total Volume, find only the max values in column "Total Stock Volume"
    
    
    ' 'ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table))'
    'This formula above is used to find the maximum value in the range 'K2:K' where the last row is specified by the variable Summary_Table on the ws
    'and then assigning that maximum value to the Cell(i, 11) of the same worksheet.
    
    
    'Gather Ticker_Names and their corresponding values for the Percent_Change & Total_Stock_Volume

        For i = 2 To Summary_Table
        
            'Calculate Maximum Percent_Change and populate results in destination ws.Cells
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Calculate Minimum Percent_Change and populate results in destination ws.Cells
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Calculate Maximum Total_Stock_Volume and populate results in destination ws.Cells
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
End Sub

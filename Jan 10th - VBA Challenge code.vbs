Sub Stock_Analysis()
    
    'Assign code to each worksheet
    For Each ws In Worksheets
    
    'Declare variables
    Dim Ticker_Name As String
    Dim Ticker As Integer
    Dim J As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    
    'Assign Value to Ticker and J
    Ticker = 2
    J = 2
    
    'Determine Lastrow of Dataset
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Label the columns for Ticker, Yearly Change, Percent Change, Total Stock Volume
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
        'Create Loop for Tickers
        For i = 2 To Lastrow
    
        'Set if conditions
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Name = ws.Cells(i, 1).Value
        
        'Assign Value to Ticker Name
        ws.Range("I" & Ticker).Value = Ticker_Name
        
        Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(J, 3).Value
        
        'Assign Value to Yearly Change
        ws.Range("J" & Ticker).Value = Yearly_Change
        
            'Set a condition for the Yearly Change where Green is positive
            If ws.Cells(Ticker, 10).Value > 0 Then
            ws.Cells(Ticker, 10).Interior.ColorIndex = 4
            
            
            'Set a condition for the Yearly Change where Red is negative
            ElseIf ws.Cells(Ticker, 10).Value < 0 Then
            ws.Cells(Ticker, 10).Interior.ColorIndex = 3
        
            End If
        
        'Calculate Percent Change
        Percent_Change = (Yearly_Change / ws.Cells(J, 3).Value)
        
        'Calculate Total Stock Volume
        Total_Stock_Volume = WorksheetFunction.Sum(Range(ws.Cells(J, 7), ws.Cells(i, 7)))
        
       'Assign Value to Percent Change, and Total Stock Volume
        ws.Range("K" & Ticker).Value = Format(Percent_Change, "Percent")
        ws.Range("L" & Ticker).Value = Total_Stock_Volume
       
       
        'Increment value of Ticker
        Ticker = Ticker + 1
        
        'Increment value of J
        J = i + 1
        
        
        End If

    
        Next i
    
    
    'Declare Variables
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As Double
    
    'Assign Initial Values to Greatest Increase, Decrease and Total Volume
    Greatest_Increase = Range("K2").Value
    Greatest_Decrease = Range("K2").Value
    Greatest_Total_Volume = Range("L2").Value
    
    
    'Determine Lastrow of Unique Tickers
    Lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Assign Value to Greatest % increase, decrease, Total Volume, Ticker, and Value
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
        
        'loop through Unique Tickers
        For i = 2 To Lastrow2
        
        
        'Set a condition to determine Greatest Increase based on percent change
        If ws.Cells(i, 11).Value > Greatest_Increase Then
        Greatest_Increase = ws.Cells(i, 11).Value
        
        ws.Range("P2").Value = ws.Cells(i, 9).Value
        
        
        End If
        
        'Retrieve greatest increase and format to percent
        ws.Range("Q2").Value = Format(Greatest_Increase, "Percent")
        
        
        'Set a condition to determine Greatest Derease based on percent change
        If ws.Cells(i, 11).Value < Greatest_Decrease Then
        Greatest_Decrease = ws.Cells(i, 11).Value
        
        ws.Range("P3").Value = ws.Cells(i, 9).Value
        
        
        End If
        
        'Retrieve greatest derease and format to percent
        ws.Range("Q3").Value = Format(Greatest_Decrease, "Percent")
        
        'Set a condition to determine Greatest total volume based on total stock volume
        If ws.Cells(i, 12).Value > Greatest_Total_Volume Then
        Greatest_Total_Volume = ws.Cells(i, 12).Value
        
        ws.Range("P4").Value = ws.Cells(i, 9).Value
        
        End If
        
        'Retrieve greatest total volume and format to scientific
        ws.Range("Q4").Value = Format(Greatest_Total_Volume, "Scientific")
        
        
        Next i

        'Format active columns to autofit
        ws.Columns("A:Q").AutoFit
        
    
    'loop through next worksheet
    Next ws
    

End Sub

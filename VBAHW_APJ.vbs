Attribute VB_Name = "Module1"
'VBA stock'

Sub StocksVBAHW():



'Defining the variables I need'
'Use Dim to assign each of them as a Variable'
Dim ticker As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease As Double
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume As Double
Dim greatest_stock_volume_ticker As String


'Loop through the 3 sheets 2016, 2015, 2014'
For Each ws In Worksheets



    'Activate all 3 sheets'
    ws.Activate

    'Last Row'
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    

    'Header for each row'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'start points for each row'
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    
    
    'Row by Row through the ticker Row'
    For i = 2 To lastRowState

    'Each ticker calculated from the whole list'
        ticker = Cells(i, 1).Value
        
        'Price at beggining of the year for each company'
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        'Sum volume of stocks for each company'
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        
        
        'New company ticker symbol run this'
        If Cells(i + 1, 1).Value <> ticker Then
            'When a new ticker symbol comes increment'
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            'Year close price for particular company'
            closing_price = Cells(i, 6)
            
            'Change in stock price for company for whole year'
            yearly_change = closing_price - opening_price
            
            'Change in price over year'
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            
            
            
        'Color Cell Red or Green for Positive or Negative Gain over the year'
            
            
            'If change value is positive cell is green'
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            'If change is negative cell is red'
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            'If change is 0 cell is yellow'
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            'Change over year shown as a percentage'
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            'change in value of each ticker shown as a percentage'
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
           
            'Set opening price to zero when a new company ticker is reached'
            opening_price = 0
            
            ' Sum total stock volume value for each company to the corresponding cell in each sheet'
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Set total stock volume back to 0 when we reach a new company ticker symbol'
            total_stock_volume = 0
        End If
        
        
        
        
        
    Next i
    
    'Challenges for HW return stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"'
    
    
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Last Row'
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'values of variables in the first row'
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    'Row by Row through the whole list of company ticker symbols'
    For i = 2 To lastRowState
    
        'Greatest increase company ticker price'
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        'Largest decrease company ticker price'
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        'Company with the most volume of ticker shares'
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Show the values for company with greatest percent increase, decrease, and highest volume of ticker shares'
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws


End Sub

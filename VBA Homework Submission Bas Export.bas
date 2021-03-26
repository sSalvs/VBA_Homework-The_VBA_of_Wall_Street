Attribute VB_Name = "Module1"
Sub StockData()


' Script to loop through sheets and count ticker values


    'Add Definitions
    
    Dim wb As Workbook
    Dim ws As Worksheet
      
      
    Set wb = ActiveWorkbook
    Set ws = Worksheets("2014")
    Set ws2 = Worksheets("2015")
    Set ws3 = Worksheets("2016")
    
       For Each ws In Worksheets
      
    'label the columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
    'define exisiting columns
    
    Dim ticker As String
    Dim Year As Date
    Dim openPrice As Double
    Dim highPrice As Double
    Dim lowPrice As Double
    Dim closePrice As Double
    Dim vol As Double
    Dim Max As Double
    Dim Min As Double
    Dim GTV As Double
    'vol = 0
    
    'define new columns
    Dim table_results As Integer
    table_results = 2
    sd = 2
    
                
        'Set up Loop
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
                          
            'Look for change in ticker name
          
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                     
                    'Take note of the ticker
                    ticker = ws.Cells(i, 1).Value
                    
                    'calculate the total stock colume per ticker
                     vol = vol + ws.Cells(i, 7).Value
                            
                    'Add open and close Price
                    openPrice = ws.Cells(sd, 3).Value
                    closePrice = ws.Cells(i, 6).Value
                                                        
                    'Add the Ticker symbol
                    ws.Range("I" & table_results).Value = ticker
                    
                    'Add the total volume
                    ws.Range("L" & table_results).Value = vol
                    
                    'define yearly change
                    yearlyChange = closePrice - openPrice
                    
                    'Add Yearly Price
                    ws.Range("J" & table_results).Value = yearlyChange
               
                     
                    'Add the percentage change column
        
                    If openPrice = 0 Then
                    percentChange = Null
                    Else
                    percentChange = (yearlyChange / openPrice)
                    End If
                  
                    ws.Range("K" & table_results).Value = percentChange
                    
                    
                    'conditional Formatting- set criteria for colours
                    If ws.Cells(table_results, 10).Value > 0 Then
                        ws.Cells(table_results, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(table_results, 10).Interior.ColorIndex = 3
                    End If
                                                                    
                    'Reset the ticker
                    vol = 0
                    sd = i + 1
                                                     
                    'count second table results
                    If table_results = 2 Then
                        Max = ws.Cells(table_results, 11)
                        Min = ws.Cells(table_results, 11)
                        GTV = ws.Range("L" & table_results).Value
                        NewTickerMax = ws.Cells(table_results, 9)
                        NewTickerMin = ws.Cells(table_results, 9)
                        NewTickerGTV = ws.Cells(table_results, 9)
                    End If
            
                    'Set If for Greatest Increase
                    If ws.Cells(table_results, 11).Value > Max Then
                        'capture new Max
                        Max = ws.Cells(table_results, 11).Value
                        NewTickerMax = ws.Cells(table_results, 9).Value
                    ElseIf Cells(table_results, 11).Value < Min Then
                        'capture new Min
                        Min = ws.Cells(table_results, 11).Value
                        NewTickerMin = ws.Cells(table_results, 9).Value
                    Else
                        'Do nothing
                        Max = Max
                        Min = Min
                    End If
                    
                    'Capture new Greatest Total Volume
                    If ws.Cells(table_results, 12).Value > GTV Then
                        'capture new gtv
                        GTV = ws.Cells(table_results, 12).Value
                        NewTickerGTV = ws.Cells(table_results, 9).Value
                    Else
                        'Do nothing
                        GTV = GTV
                    End If
                      
                    'Add new values to table
                    ws.Cells(2, 17).Value = Max
                    ws.Cells(2, 16).Value = NewTickerMax
                    ws.Cells(3, 17).Value = Min
                    ws.Cells(3, 16).Value = NewTickerMin
                    ws.Cells(4, 17).Value = GTV
                    ws.Cells(4, 16).Value = NewTickerGTV
                    
                    'Keep adding to table
                    table_results = table_results + 1
       
            Else
                'Add Total Stock Volume
                vol = vol + ws.Cells(i, 7).Value
            '    Cells(i, 7).Select
           End If
                     
        Next i
    
    'Format cells
    Columns("K").NumberFormat = "0.00%"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"

Next ws

End Sub


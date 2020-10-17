Attribute VB_Name = "Module1"
Sub alphabetical_testing():

'Identify variables
 Dim ws As Worksheet
 Dim max_change As Double
 Dim min_change As Double
 Dim max_ticker As String
 Dim min_ticker As String
 Dim greatest_stvolume As LongLong
 Dim greatest_stvticker As String

 max_change = 0
 min_change = 0
 greatest_volume = 0
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
'Insert & Format Headers: "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume"
    'Ticker Header1:
        Range("A1") = "Ticker"
        Range("A1").Font.Bold = True
        Range("A1").Font.ColorIndex = 1
        Range("A1").HorizontalAlignment = xlCenter
        
    'Date Header
        Range("B1") = "Date"
        Range("B1").Font.Bold = True
        Range("B1").Font.ColorIndex = 1
        Range("B1").HorizontalAlignment = xlCenter
    
    'Open Header
        Range("C1") = "Open"
        Range("C1").Font.Bold = True
        Range("C1").Font.ColorIndex = 1
        Range("C1").HorizontalAlignment = xlCenter
        
    'High Header
        Range("D1") = "High"
        Range("D1").Font.Bold = True
        Range("D1").Font.ColorIndex = 1
        Range("D1").HorizontalAlignment = xlCenter
    'Low Header
        Range("E1") = "Low"
        Range("E1").Font.Bold = True
        Range("E1").Font.ColorIndex = 1
        Range("E1").HorizontalAlignment = xlCenter
        
    'Close Header
        Range("F1") = "Close"
        Range("F1").Font.Bold = True
        Range("F1").Font.ColorIndex = 1
        Range("F1").HorizontalAlignment = xlCenter
        
    'Volumn Header
        Range("G1") = "Volume"
        Range("G1").Font.Bold = True
        Range("G1").Font.ColorIndex = 1
        Range("G1").HorizontalAlignment = xlCenter
        
'Summary Headers:
        'Ticker Header2
        Range("I1") = "Ticker"
        Range("I1").Font.Bold = True
        Range("I1").Font.ColorIndex = 1
        Range("I1").HorizontalAlignment = xlCenter
        
        'Yearly Change
        Range("J1") = "Yearly Change"
        Range("J1").Font.Bold = True
        Range("J1").Font.ColorIndex = 1
        Range("J1").HorizontalAlignment = xlCenter
        
        'Percent Change
        Range("K1") = "Percent Change"
        Range("K1").Font.Bold = True
        Range("K1").Font.ColorIndex = 1
        Range("K1").HorizontalAlignment = xlCenter
        
        
        'Total Stock Volume
        Range("L1") = "Total Stock Volume"
        Range("L1").Font.Bold = True
        Range("L1").Font.ColorIndex = 1
        Range("L1").HorizontalAlignment = xlCenter
  
    
    'Items in "Ticker" Column: Ticker Names listed aphabetically
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim pct_change As Double
        Dim total_volume As LongLong
        Dim summary_i As Integer
        
        
        ticker = Range("A" & 2)
        open_price = Range("C" & 2)
        total_volume = 0
        summary_i = 2
        
        For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
            'msgBox(ticker)
            'Range("I" & i) = ticker
            total_volume = total_volume + Range("G" & i)
            
            'check if cell under is different
            If Range("A" & i + 1) <> ticker Then
                
                'Need to start tracking a new ticker
                close_price = Range("F" & i)
                'MsgBox ("Ticker Changed at " & i)
                
                'Calculate summary
                yearly_change = close_price - open_price
                
                If open_price <> 0 Then
                    pct_change = yearly_change / open_price
                    
                Else
                    pct_change = 0
                    
                    
                End If
                
                'Insert summary for previous ticker
                Range("I" & summary_i) = ticker
                
                'Insert "Yearly Change" calculations
                Range("J" & summary_i) = yearly_change
                
                'Conditional Formating:
                
                'Highlight positive change in green
                If Range("J" & summary_i) >= 0 Then
                    Range("J" & summary_i).Interior.ColorIndex = 4
    
                    'Highlight negative change in red.
                    ElseIf Range("J" & summary_i) < 0 Then
                    Range("J" & summary_i).Interior.ColorIndex = 3
    
                End If
                
                'Insert "Percent Change" Calculations
                Range("K" & summary_i) = pct_change
                Range("K" & summary_i).NumberFormat = "0.00%"
                
                'Identify Greatest % Increase
                If Range("K" & summary_i) > max_change Then
                        max_change = Range("K" & summary_i)
                        max_ticker = Range("I" & summary_i)
                
                ElseIf Range("K" & summary_i) < min_change Then
                        min_change = Range("K" & summary_i)
                        min_ticker = Range("I" & summary_i)
                
                End If
                
                'Insert "Total Volume" Calculations
                Range("L" & summary_i) = total_volume
                
                    'Identify Greatest Total Volume
                     If Range("L" & summary_i) > greatest_stvolume Then
                        greatest_stvolume = Range("L" & summary_i)
                        greatest_stvticker = Range("I" & summary_i)
                
                    End If
                
                'Reset initial values
                ticker = Range("A" & i + 1)
                
                open_price = Range("C" & i + 1)
                
                summary_i = summary_i + 1
                
                total_volume = 0
            
            End If
        
        Next i
        
        'Challenge Headers and Row Labels:
        
        'Ticker Header3
        Range("S4") = "Ticker"
        Range("S4").Font.Bold = True
        Range("S4").Font.ColorIndex = 1
        Range("S4").HorizontalAlignment = xlCenter
        
        'Value Header
        Range("T4") = "Value"
        Range("T4").Font.Bold = True
        Range("T4").Font.ColorIndex = 1
        Range("T4").HorizontalAlignment = xlCenter
        
        'Greatest % Increase
        Range("R5") = "Greatest % Increase"
        Range("S5") = max_ticker
        Range("T5") = max_change
        Range("T5").NumberFormat = "0.00%"
        
        'Greatest % Decrease
        Range("R6") = "Greatest % Decrease"
        Range("S6") = min_ticker
        Range("T6") = min_change
        Range("T6").NumberFormat = "0.00%"
        
        'Greatest Total Volume
        Range("R7") = "Greatest Total Volume"
        Range("S7") = greatest_stvticker
        Range("T7") = greatest_stvolume
        Range("T7").NumberFormat = "0"
    
            
    Next ws
        'Copy Data from Sheet "P" (source) to Sheet "A" (destination)
        Sheets("P").Range("R4:T7").Copy Destination:=Sheets("A").Range("N1")

End Sub

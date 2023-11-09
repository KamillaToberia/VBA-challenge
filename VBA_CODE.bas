Attribute VB_Name = "Module1"
Sub VBAChallenge()

'Loop through all the sheets
For Each ws In Worksheets
Dim worksheetName As String
Dim startPrice As Double
Dim closePrice As Double
Dim volTotal As Double
Dim pcTotal As Double
Dim Ticker As Long
Dim LastRow As Long
Dim i As Long
Dim yearlyChange As Double
Dim TickerMax As String
Dim TickerMin As String
Dim MaxPc As Double
Dim MinPc As Double
Dim MaxVolTicker As String
Dim MaxVol As Double
worksheetName = ws.Name
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Ticker = 2
startPrice = 0
closePrice = 0
volTotal = 0
yearlyChange = 0
pcTotal = 0
TickerMax = ""
TickerMin = ""

MaxPc = 0
MinPc = 0
MaxVolTicker = ""
MaxVol = 0

    'Create the columns and name it
    ws.Range("I1").EntireColumn.Insert
    ws.Range("J1").EntireColumn.Insert
    ws.Range("K1").EntireColumn.Insert
    ws.Range("L1").EntireColumn.Insert
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest% Increase"
    ws.Cells(3, 15).Value = "Greatest% Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    startPrice = ws.Cells(2, 3).Value
    'Loop
    For i = 2 To LastRow
    
    
    'Find the unique value for the Ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
    
    'Finding the Yearly Change
        closePrice = ws.Cells(i, 6).Value
        yearlyChange = closePrice - startPrice
    
    'Find the Percent Change
        If startPrice <> 0 Then
        pcTotal = yearlyChange / startPrice
        ws.Cells(Ticker, 11).NumberFormat = "0.00%"
        End If
    'Find the Total Stock Volume
    
        volTotal = volTotal + ws.Cells(i, 7).Value
    'Alocate the Yearly Change,Volume Total and percent total
        ws.Cells(Ticker, 10).Value = yearlyChange
   
        ws.Cells(Ticker, 11).Value = pcTotal
    
        ws.Cells(Ticker, 12).Value = volTotal
        
        'Conditional Format
        If ws.Cells(Ticker, 10).Value >= 0 Then
        ws.Cells(Ticker, 10).Interior.ColorIndex = 4
        ws.Cells(Ticker, 11).Interior.ColorIndex = 4
    
        ElseIf ws.Cells(Ticker, 10).Value < 0 Then
        ws.Cells(Ticker, 10).Interior.ColorIndex = 3
        ws.Cells(Ticker, 11).Interior.ColorIndex = 3
        End If
        
        
        'Reset variables
        Ticker = Ticker + 1
        closePrice = 0
        yearlyChange = 0
        startPrice = ws.Cells(i + 1, 3).Value
   
   
         If (pcTotal > MaxPc) Then
        MaxPc = pcTotal
        TickerMax = ws.Cells(i, 1).Value
        ws.Cells(2, 16).Value = TickerMax
        ws.Cells(2, 17).Value = MaxPc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ElseIf (pcTotal < MinPc) Then
        MinPc = pcTotal
        TickerMin = ws.Cells(i, 1).Value
        ws.Cells(3, 16).Value = TickerMin
        ws.Cells(3, 17).Value = MinPc
        ws.Cells(3, 17).NumberFormat = "0.00%"
        End If
                           
        If (volTotal > MaxVol) Then
        MaxVol = volTotal
        MaxVolTicker = ws.Cells(i, 1).Value
        ws.Cells(4, 16).Value = MaxVolTicker
        ws.Cells(4, 17).Value = MaxVol
        End If
                    
        pcTotal = 0
        volTotal = 0
        
        Else: volTotal = volTotal + ws.Cells(i, 7).Value
        End If
   
    Next i
    Worksheets(worksheetName).Columns("A:Z").AutoFit

    Next ws

    End Sub



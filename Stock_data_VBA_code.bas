Attribute VB_Name = "Module1"
Sub wall_street()

'1.Create a script that will loop through all the stocks for one year and output the following information.
'   - Output the ticker symbol
'   - Output the yearly change from open tp close
'   - Output the percent change from open to close
'   - Output the total stock volume of the stock

'---------------------------------------------------------------------
'Create a forloop to make sure these rules are being spread across all worksheets
    For Each ws In Worksheets

'---------------------------------------------------------------------
        'Add in and set variables for new information you will be calculating
        Dim ticker As String
        Dim ticker_counter As Double
        Dim yearly_change As Double
        Dim yearly_open As Double
        Dim yearly_close As Double
        Dim percent_change As Double
        Dim total_volume As Double
        'set ticket counter to 2
        Dim ticket_counter As Long
        ticker_counter = 2
        Dim previous_amount As Long
        previous_amount = 2
        Dim holder As Long
        holder = 2
        Dim holder2 As Long
        holder2 = 2
    '---------------------------------------------------------------------
        'Create new columns to track new data being calculated
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
       
    '---------------------------------------------------------------------
        'Set last row and collect value
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
        'set total volume to 0 so that it counts each value when it reads the row
        total_volume = 0
    '---------------------------------------------------------------------
        'Set for loop to run through the second row (i = 2) to the last row including all columns
        For i = 2 To LastRow
           
            'Calculate the value of total stock volume
            total_volume = total_volume + ws.Cells(i, 7).Value
           
            'Locate the values for Tickers
            ticker = ws.Cells(i, 1).Value
            yearly_open = ws.Cells(previous_amount, 3)
           
            'Check next row to see if ticker is the same or if it has changed value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearly_close = ws.Cells(i, 6)
               
                'Add the value of the ticker and the balue for yearly change in the correct cells
                ws.Cells(ticker_counter, 10).Value = yearly_close - yearly_open
                ws.Cells(ticker_counter, 9).Value = ticker
           
                'Putting previous calculation for total volume into new column
                ws.Cells(ticker_counter, 12).Value = total_volume


                ' Determine Percent Change
                If yearly_open = 0 Then
                ws.Cells(ticker_counter, 11).Value = Null
                Else
                ws.Cells(ticker_counter, 11).Value = (yearly_close - yearly_open) / yearly_open
                End If
               
                ws.Cells(ticker_counter, 11).NumberFormat = "0.00%"
   
                'Add Conditional Formatting -
                'If Yearly Change is greater than 0 then Green
                If ws.Cells(ticker_counter, 10).Value > 0 Then
                ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
               
                'If Yearly Change is less than 0 then Red
                Else
                ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
                End If
               
                'Reset loop and move along to next row
                total_volume = 0
                ticker_counter = ticker_counter + 1
                holder = holder + 1
                previous_amount = i + 1
            End If
        Next i
        'make all columns fit the content
        Columns("J").AutoFit
        Columns("K").AutoFit
        Columns("L").AutoFit
    Next ws
End Sub


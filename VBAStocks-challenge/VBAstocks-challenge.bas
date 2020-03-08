Attribute VB_Name = "Module1"

Sub stocks()

'---------------------------------
'Hans Engelbrecht
'VBA - Challenge
'Stocks w/ challenge questions
'---------------------------------


For Each ws In Worksheets

 ' Set an initial variable for holding the brand name
  Dim ticker As String
  
  'Set row counter
  Dim row_counter As Integer
  row_counter = 1

  ' Set an initial variable for holding the volume
  Dim volume As Double
  volume = 0
  
  ' Set variable for starting price and ending price
  Dim op As Double
  Dim ep As Double
  
  'Set variable for the price difference
  Dim pri_diff As Double

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
     
  'Variables for Maximum percent increase, maximum percent decrease, Max volume, and the associated tickers
    Dim max_val As Double
    Dim min_val As Double
    Dim max_vol As LongLong
    Dim max_tic As String
    Dim min_tic As String
    
    max_val = 0
    min_val = 0
    max_vol = 0
  
    'Add Headers
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Year Open Price"
  ws.Range("K1").Value = "Year Close Price"
  ws.Range("L1").Value = "Change Price Per share"
  ws.Range("M1").Value = "Change percent"
  ws.Range("N1").Value = "Total Volume"
  
    'Determine last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all tickers
    For i = 2 To LastRow
    
        ' Check if we are still within the same ticker name, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ' Set the ticker name.
            ticker = ws.Cells(i, 1).Value
        
            ' Add to the volume total.
            volume = volume + ws.Cells(i, 7).Value
            
            'Print the ticker in the summary table
            ws.Range("I" & Summary_Table_Row).Value = ticker
        
            'Print the total volume to the summary table
            ws.Range("N" & Summary_Table_Row).Value = volume
            
            'Print the opening price to the summary table
            ws.Range("J" & Summary_Table_Row).Value = op
            
            'Print the closing price to the summary table
            ws.Range("K" & Summary_Table_Row).Value = ws.Cells(i, 6).Value
            
            'Calculate the price difference
            ws.Range("L" & Summary_Table_Row).Value = ws.Range("K" & Summary_Table_Row).Value - ws.Range("J" & Summary_Table_Row).Value
            ws.Range("L" & Summary_Table_Row).NumberFormat = "$#.##"
                If ws.Range("L" & Summary_Table_Row).Value < 0 Then
                 ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                     
                Else
                    ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                        
                End If
            
            'Determine there are any 0s in cells that will be divided.
            If ws.Range("L" & Summary_Table_Row).Value = 0 Or ws.Range("j" & Summary_Table_Row).Value = 0 Then
                ws.Range("M" & Summary_Table_Row).Value = 0
            Else
                'Calculate the price difference percent/Convert number format to percentage/apply color to positive and negitive change
                ws.Range("M" & Summary_Table_Row).Value = ws.Range("L" & Summary_Table_Row).Value / ws.Range("j" & Summary_Table_Row).Value
                ws.Range("M" & Summary_Table_Row).NumberFormat = "0.0%"
                If ws.Range("M" & Summary_Table_Row).Value < 0 Then
                    ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 3
                           
                Else
                    ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 4
                           
                End If
             End If
             
            If ws.Range("M" & Summary_Table_Row).Value > max_val Then
                max_val = ws.Range("M" & Summary_Table_Row).Value
                max_tic = ws.Range("I" & Summary_Table_Row).Value
            End If
            If ws.Range("M" & Summary_Table_Row).Value < min_val Then
                min_val = ws.Range("M" & Summary_Table_Row).Value
                min_tic = ws.Range("I" & Summary_Table_Row).Value
            End If
            
             
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
        
            ' Reset volume to zero
            volume = 0
            ' Reset row counter to 1
            row_counter = 1
            
        'If the cell immediately following a row is the same ticker...
        Else
            
            'Add to the total volume and get max volume
            volume = volume + ws.Cells(i, 7).Value
            If volume > max_vol Then
                max_vol = volume
                vol_tic = ws.Cells(i, 1).Value
            End If
            
            If row_counter = 1 Then
            op = ws.Cells(i, 3).Value
            End If
            
            'reset row counter
            row_counter = row_counter + 1
            
        End If
  
    
    Next i



'Determine greatest % increase, greatest % decrease, and greatest total volume with loops
    'LastRow = Cells(Rows.Count, 13).End(xlDown).Row


    
    
    'Add Challenge data, including headers
  ws.Range("Q1").Value = "Ticker"
  ws.Range("R1").Value = "Value"
  ws.Range("P2").Value = "Largest % Gain"
  ws.Range("P3").Value = "Larges % Loss"
  ws.Range("P4").Value = "Largest Volume"
  ws.Range("Q2").Value = max_tic
  ws.Range("R2").NumberFormat = "0.0%"
  ws.Range("Q3").Value = min_tic
  ws.Range("R3").NumberFormat = "0.0%"
  ws.Range("Q4").Value = vol_tic
  ws.Range("R2").Value = max_val
  ws.Range("R3").Value = min_val
  ws.Range("R4").Value = max_vol
  
  
  
Next ws

End Sub

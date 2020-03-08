Attribute VB_Name = "Module1"
Sub stocks_required()



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
  
    'Add Headers
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Year Open Price"
  Range("K1").Value = "Year Close Price"
  Range("L1").Value = "Change Price Per share"
  Range("M1").Value = "Change percent"
  Range("N1").Value = "Total Volume"
  
    'Determine last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all tickers
    For i = 2 To LastRow
    
        ' Check if we are still within the same ticker name, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            ' Set the ticker name.
            ticker = Cells(i, 1).Value
        
            ' Add to the volume total.
            volume = volume + Cells(i, 7).Value
            
            'Print the ticker in the summary table
            Range("I" & Summary_Table_Row).Value = ticker
        
            'Print the total volume to the summary table
            Range("N" & Summary_Table_Row).Value = volume
            
            'Print the opening price to the summary table
            Range("J" & Summary_Table_Row).Value = op
            
            'Print the closing price to the summary table
            Range("K" & Summary_Table_Row).Value = Cells(i, 6).Value
            
            'Calculate the price difference
            Range("L" & Summary_Table_Row).Value = Range("K" & Summary_Table_Row).Value - Range("J" & Summary_Table_Row).Value
            Range("L" & Summary_Table_Row).NumberFormat = "$#.##"
                If Range("L" & Summary_Table_Row).Value < 0 Then
                    Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
            
            'Determine there are any 0s in cells that will be divided.
            If Range("L" & Summary_Table_Row).Value = 0 Or Range("j" & Summary_Table_Row).Value = 0 Then
                Range("M" & Summary_Table_Row).Value = 0
            Else
                'Calculate the price difference percent/Convert number format to percentage/apply color to positive and negitive change
                Range("M" & Summary_Table_Row).Value = Range("L" & Summary_Table_Row).Value / Range("j" & Summary_Table_Row).Value
                Range("M" & Summary_Table_Row).NumberFormat = "0.0%"
                If Range("M" & Summary_Table_Row).Value < 0 Then
                    Range("M" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    Range("M" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
             End If
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
        
            ' Reset volume to zero
            volume = 0
            ' Reset row counter to 1
            row_counter = 1
            
        'If the cell immediately following a row is the same ticker...
        Else
            
            'Add to the total volume
            volume = volume + Cells(i, 7).Value
            
            If row_counter = 1 Then
            op = Cells(i, 3).Value
            End If
            
            'reset row counter
            row_counter = row_counter + 1
            
        End If
    Next i


End Sub



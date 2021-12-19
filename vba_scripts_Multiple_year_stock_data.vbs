Sub stock_summary()

'Set lastrow
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set title rows and columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Set dimensions & set initial values
Dim ticker As String

Dim total_volume As Double
total_volume = 0

Dim yearly_change As Double
yearly_change = 0

Dim percent_change As Double
percent_change = 0

Dim back_to_open As Long
back_to_open = 0

Dim yearly_open As Double
yearly_open = 0

Dim yearly_close As Double
yearly_close = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Loop through the rows of the worksheet
  For i = 2 To LastRow

    ' Check to see if there is a match; otherwise move down
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
      'Set the values
      ticker = Cells(i, 1).Value
      
      yearly_open = Cells(i - back_to_open, 3).Value
      
      yearly_close = Cells(i, 6).Value

  
      'Calculate the values
      total_volume = total_volume + Cells(i, 7).Value
      
      yearly_change = yearly_close - yearly_open
  
        'If yearly_open is 0, then there will be overflow error
        If yearly_open = 0 And yearly_change = 0 Then
            percent_change = 0
            
        ElseIf yearly_open = 0 Then
            yearly_open = 1
            percent_change = yearly_change / yearly_open
        Else
            percent_change = yearly_change / yearly_open
        End If
  
      'print proper names to columns
      Range("I" & Summary_Table_Row).Value = ticker
      Range("J" & Summary_Table_Row).Value = yearly_change
      Range("K" & Summary_Table_Row).Value = percent_change
      Range("L" & Summary_Table_Row).Value = total_volume

      'Add proper color, 4 is green, 3 is red
      If (yearly_change >= 0) Then
      Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      End If
      
      If (yearly_change < 0) Then
      Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If

      Summary_Table_Row = Summary_Table_Row + 1
  
      'Reset values
      total_volume = 0
      
      back_to_open = 0

    Else
  
      'Accumulate the totals
      total_volume = total_volume + Cells(i, 7).Value
        
      'Add the counter
      back_to_open = back_to_open + 1

    End If

  Next i

'Set the dimensions and initiate the values
Dim max_percent As Double
max_percent = 0

Dim min_percent As Double
min_percent = 0

Dim max_total As Double
max_total = 0

  'Find the largest percentage and print the ticker and the number
  For i = 2 To LastRow

    If Cells(i, 11).Value > max_percent Then
    max_percent = Cells(i, 11).Value
    Range("P2").Value = Cells(i, 9).Value
    Range("Q2").Value = max_percent
    
    End If

  Next i

  'Find the smallest percentage and print the ticker and the number  
  For i = 2 To LastRow

    If Cells(i, 11).Value < min_percent Then
    min_percent = Cells(i, 11).Value
    Range("P3").Value = Cells(i, 9).Value
    Range("Q3").Value = min_percent
    
    End If

  Next i
  
  'Find the largest volume and print the ticker and the number
  For i = 2 To LastRow

    If Cells(i, 12).Value > max_total Then
    max_total = Cells(i, 12).Value
    Range("P4").Value = Cells(i, 9).Value
    Range("Q4").Value = max_total
    
    End If

  Next i

'Percentage formatting
Range("K2:K" & LastRow).NumberFormat = "0.00%"
Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"

'Autofit columns width
Columns("A:Q").AutoFit

End Sub

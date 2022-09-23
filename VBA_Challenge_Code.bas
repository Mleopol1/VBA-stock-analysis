Attribute VB_Name = "Module1"
Sub stock_ticker()

  ' Set variables and initial values
  Dim ticker As String
  Dim open_price As Double
  Dim close_price As Double
  Dim yearly_change As Double
  Dim percent_change As Double
  Dim total_volume As Double
  total_volume = 0
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Set initial open price for first ticker
  open_price = Cells(2, 3).Value
  
  ' Find last non-blank cell in column A
  lrow = Cells(Rows.Count, 1).End(xlUp).Row

 ' Create headers
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"
  
    For i = 2 To lrow

    ' Checking if we are still in the same stock, if not,
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set the ticker name
        ticker = Cells(i, 1).Value
      
        ' Set the closing price
        close_price = Cells(i, 6).Value
    
        ' Calculate the yearly change
        yearly_change = close_price - open_price
          
        ' Calculate the percent change
        percent_change = (yearly_change / open_price)
        
        ' Add to the total stock volume
        total_volume = total_volume + Cells(i, 7).Value
          
        ' Print the ticker name in the Summary Table
        Range("I" & Summary_Table_Row).Value = ticker
    
        ' Print the yearly change to the Summary Table
        Range("J" & Summary_Table_Row).Value = yearly_change
          
        ' Print the percent change to the Summary Table
        Range("K" & Summary_Table_Row).Value = percent_change
          
        ' Print the total stock volume to the Summary Table
        Range("L" & Summary_Table_Row).Value = total_volume
    
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the values
        open_price = Cells(i + 1, 3).Value
        yearly_change = 0
        percent_change = 0
        total_volume = 0

    ' If the cell immediately following a row is the same stock,
    Else

        ' Add to the total stock volume
        total_volume = total_volume + Cells(i, 7).Value

    End If
  
  Next i

End Sub

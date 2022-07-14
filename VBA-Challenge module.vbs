Attribute VB_Name = "Module1"
Sub Stocks():
'Create variables
    Dim Ws As Worksheet
    Dim LastRow As Long
    Dim ticker_row As Integer
    Dim ticker_name As String
    Dim open_ As Double
    Dim close_ As Double
    Dim change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim LastRow2 As Long

'Set Header for each worksheet
    For Each Ws In ThisWorkbook.Worksheets
            Ws.Cells(1, 9).Value = "Ticker"
            Ws.Cells(1, 10).Value = "Yearly Change"
            Ws.Cells(1, 11).Value = "Percent Change"
            Ws.Cells(1, 12).Value = "Total Stock Volume"

'Define Variables
        ticker_row = 2
    
        open_ = Cells(2, 3).Value
    
        LastRow = Range("A1").End(xlDown).Row
    
        LastRow2 = Range("J1").End(xlDown).Row
    
        total_volume = 0

'Conditional Loop to run through each row and grab needed data
        For i = 2 To LastRow
    
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Check if next ticker cell is the same as previous
                ticker_name = Cells(i, 1).Value
                'add total volume from current cell to previous cell
                total_volume = total_volume + Cells(i, 7).Value
            'Display ticker name in new column
                Range("I" & ticker_row).Value = ticker_name
            'subtract opening price from year end price
                change = Cells(i, 6).Value - open_
                'display yearly change in column J
                Range("J" & ticker_row).Value = change
            'Calculate percent change by dividing yearly change by yearly opening price
                percent_change = (change / open_) * 100
        'Display percent change in column K
                Range("K" & ticker_row).Value = percent_change
            'display sum of volume in column L
                Range("L" & ticker_row).Value = total_volume
            
                open_ = Cells(i + 1, 3).Value
          'move on to next row in summary table
                ticker_row = ticker_row + 1
            'reset total volume
                total_volume = 0
            Else:
            'add total volume
                total_volume = total_volume + Cells(i, 7).Value
         
            End If
        'next row
        Next i
    'conditional formatting
        For j = 2 To LastRow2
    
            If Cells(j, 10).Value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            Else
                Cells(j, 10).Interior.ColorIndex = 3
            End If
       'next row
        Next j
    'next worksheet
    Next
    
End Sub

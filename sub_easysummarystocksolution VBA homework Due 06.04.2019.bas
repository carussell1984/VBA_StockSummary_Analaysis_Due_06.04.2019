Attribute VB_Name = "Module1"
Sub easysummarystock()


    'declare all variables
    Dim ws As Worksheet
    Dim tickerName As String
    Dim sumTickerVolume As Double
    
    sumTickerVolume = 0

    For Each ws In Worksheets

    ws.Activate

    'put the other headers inplace for easy & hard homework solutions
    ws.Range("J1").Value = "Tickerstock_Symbol"
    ws.Range("K1").Value = "Sum_stock_volume"


    'format the ranges for headers
    ws.Range("J1:K1").Font.FontStyle = "Bold"
    ws.Range("J1:K1").EntireColumn.AutoFit


    'insert a way to find the last cell in dynamic column range
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row



        'prepare to loop through row in column 1 to fins name of each tickerName and then print to "J2"
         Dim nextrow As Integer
         
         
         nextrow = 2
         For i = 2 To lastrow
         
         
         sumTickerVolume = sumTickerVolume + ws.Cells(i, 7)

             If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

               'the tickerName is stored
               tickerName = ws.Cells(i, 1).Value
    
                'the tickerName is the printed in the summary column
               ws.Cells(nextrow, 10).Value = tickerName

               'sums the final tickerstock in the column
               sumTickerVolume = sumTickerVolume + ws.Cells(i, 7)
               
               'print the total sumTickerVolume for the set it was working on
               ws.Cells(nextrow, 11).Value = sumTickerVolume
               
               'resets the row to print during the next tickerstocks & resets the sumTickerVolume to 0 for the next tickerstock set
               nextrow = nextrow + 1
               
               sumTickerVolume = 0
               

               Else
               
               'this is allowing the tickerstocks with the same symbol to keep summing
               sumTickerVolume = sumTickerVolume + ws.Cells(i, 7)
               
               
               

            End If

        Next i

    Next ws

End Sub

Attribute VB_Name = "Module1"
Sub mediumsummarystock()





    'declare all variables
    Dim ws As Worksheet
    Dim tickerName As String
    Dim sumTickerVolume As Double
    Dim openStock As Double
    Dim closeStock As Double
    Dim percentChange As Double
    Dim deltaChange As Double


    sumTickerVolume = 0


    For Each ws In Worksheets

    ws.Activate


    'put the other headers inplace for easy & hard homework solutions
    ws.Range("J1").Value = "Tickerstock_Symbol"
    ws.Range("K1").Value = "Delta_Yearly_Change"
    ws.Range("L1").Value = "%_Yearly_Change"
    ws.Range("M1").Value = "Sum_stock_volume"


   'format the ranges for headers
    ws.Range("J1:M1").Font.FontStyle = "Bold"
    ws.Range("J1:M1").EntireColumn.AutoFit


   'insert a way to find the last cell in dynamic column range
    LastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row



   'prepare to loop through row in column 1 to fins name of each tickerName and then print to "J2"
    Dim nextrow As Integer

    nextrow = 2



         For i = 2 To LastRow

            sumTickerVolume = sumTickerVolume + ws.Cells(i, 7)

             If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

                openStock = ws.Cells(i, 3).Value

             ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

               'get close stock price from

               closeStock = ws.Cells(i, 6).Value

               deltaChange = closeStock - openStock

               ws.Cells(nextrow, 11).Value = deltaChange

                   If openStock = 0 Then

                      percentChange = 0

                   Else

                       percentChange = deltaChange / openStock
                   
                   End If

                ws.Cells(nextrow, 12).Value = percentChange

               'the tickerName is stored
               tickerName = ws.Cells(i, 1).Value

               'the tickerName is the printed in the summary column
               ws.Cells(nextrow, 10).Value = tickerName

               'sums the final tickerstock in the column
               sumTickerVolume = sumTickerVolume + ws.Cells(i, 7)

               'print the total sumTickerVolume for the set it was working on
               ws.Cells(nextrow, 13).Value = sumTickerVolume

               'resets the row to print during the next tickerstocks & resets the sumTickerVolume to 0 for the next tickerstock set
               nextrow = nextrow + 1

               sumTickerVolume = 0

               Else

               'this is allowing the tickerstocks with the same symbol to keep summing
               sumTickerVolume = sumTickerVolume + ws.Cells(i, 7)


            End If

        Next i

        ws.Columns("L").NumberFormat = "0.00%"

        'findlast rows for deltachange column
        deltaChangeLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        

        For x = 2 To deltaChangeLastRow

            If ws.Cells(x, 11).Value > 0 Then

                ws.Cells(x, 11).Interior.ColorIndex = 4

            'fills cells with green
            ElseIf ws.Cells(x, 11).Value <= 0 Then

            'fill cells with red
                ws.Cells(x, 11).Interior.ColorIndex = 3

            End If

        Next x
        
'Begin Hard Solution

'print the table columns and row values
ws.Range("P2").Value = "Greatest_%_Increase"
ws.Range("P3").Value = "Greatest_%_Decrease"
ws.Range("P4").Value = "Greastest_Total_Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"

'Delcare values
Dim maxIncrease As Double
Dim minIncrease As Double
Dim greatestSum As Double
Dim summaryTicker As String
Dim summaryTicker2 As String
Dim summaryTicker3 As String

Dim z As Long
 


'set values = to something for the if then statement
maxIncrease = ws.Cells(2, 12)
minIncrease = ws.Cells(2, 12)
greatestSum = ws.Cells(2, 13)

'find last row in summary columns to loop through cells
lastSummaryRow = ws.Cells(Rows.Count, 12).End(xlUp).Row


For z = 2 To lastSummaryRow

    If ws.Cells(z, 12).Value >= maxIncrease Then
    
       maxIncrease = ws.Cells(z, 12).Value
       summaryTicker = ws.Cells(z, 10).Value
    
    End If
    
    If ws.Cells(z, 12).Value <= minIncrease Then
    
       minIncrease = ws.Cells(z, 12).Value
       summaryTicker2 = ws.Cells(z, 10).Value
    
    End If
    
    If ws.Cells(z, 13).Value >= greatestSum Then
    
       greatestSum = ws.Cells(z, 13).Value
       summaryTicker3 = ws.Cells(z, 10).Value
    
    End If

Next z

ws.Cells(2, 18).Value = maxIncrease
ws.Cells(2, 17).Value = summaryTicker

ws.Cells(3, 18).Value = minIncrease
ws.Cells(3, 17).Value = summaryTicker2

ws.Cells(4, 18).Value = greatestSum
ws.Cells(4, 17).Value = summaryTicker3

'format cells:

'Bold the row & column titles
ws.Range("P2:P4").Font.FontStyle = "Bold"
ws.Range("Q1:R1").Font.FontStyle = "Bold"

'Format the max/min values as percentages
ws.Range("R2:R3").NumberFormat = "0.00%"

'format the ranges to autfit
ws.Range("P1:R4").EntireColumn.AutoFit


    





    Next ws

End Sub

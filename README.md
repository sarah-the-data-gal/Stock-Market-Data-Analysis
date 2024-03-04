# Data Stock Analysis
Sub dataStockVolume():
    'Loop through all worksheets
    For Each ws In Worksheets
' holding the ticker name and total stock volume
        Dim ticker_name As String

' open price and close price
        Dim Opening_Price As Double
        Dim Closing_Price As Double

'yearly change and percent change
        Dim Yearly_Change_Value As Double
        Dim Percentage_Change As Double

' greatest % increase, greatest % decrease, greatest total volume and their ticker names
        Dim Max_Percent_Increase As Double
        Dim Max_Percent_Decrease As Double
        Dim Max_Total_Volume As Double
        Dim Max_Percent_Increase_Ticker As String
        Dim Max_Percent_Decrease_Ticker As String
        Dim Max_Total_Volume_Ticker As String

        'set initial variable for date's open price
        Dim Row_Index As Long
        Row_Index = 2
        Total_Stock_Volume = 0

 'location for different names of stocks
        Dim Summary_Row As Integer
        Summary_Row = 2

'Header names
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

 'Lastrow
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 To Last_Row:
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ticker_name = ws.Cells(i, 1).Value

                Total_Stock_Volume = Total_Stock_Volume + ws.Range("G" & i).Value

                ws.Range("I" & Summary_Row).Value = ticker_name

 'Print total stock volume
                ws.Range("L" & Summary_Row).Value = Total_Stock_Volume

 'Calculate yearly change and percent change
                Opening_Price = ws.Range("C" & Row_Index).Value
                Closing_Price = ws.Range("F" & i).Value
                Yearly_Change_Value = Closing_Price - Opening_Price

                If Opening_Price = 0 Then
                    Percentage_Change = 0
                Else
                    Percentage_Change = Yearly_Change_Value / Opening_Price
                End If

'Print values of yearly change and percent change
                ws.Range("J" & Summary_Row).Value = Yearly_Change_Value
                ws.Range("K" & Summary_Row).Value = Percentage_Change
                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"

'Conditional formatting  highlight positive change in green and negative change in red
                If ws.Range("J" & Summary_Row).Value > 0 Then
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                End If

                'add one to the summary table row
                Summary_Row = Summary_Row + 1
                Row_Index = i + 1

                'reset the total stock volume
                Total_Stock_Volume = 0
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Range("G" & i).Value

            End If

        Next i

 'set the first ticker's percent change and total stock volume as the greatest ones
        Max_Percent_Increase = ws.Range("K2").Value
        Max_Percent_Decrease = ws.Range("K2").Value
        Max_Total_Volume = ws.Range("L2").Value

 'Define last row of Ticker column
        Lastrow_Ticker = ws.Cells(Rows.Count, "I").End(xlUp).Row

 'Loop through each row of Ticker column to find the greatest results
        For r = 2 To Lastrow_Ticker:
            If ws.Range("K" & r + 1).Value > Max_Percent_Increase Then
                Max_Percent_Increase = ws.Range("K" & r + 1).Value
                Max_Percent_Increase_Ticker = ws.Range("I" & r + 1).Value
            ElseIf ws.Range("K" & r + 1).Value < Max_Percent_Decrease Then
                Max_Percent_Decrease = ws.Range("K" & r + 1).Value
                Max_Percent_Decrease_Ticker = ws.Range("I" & r + 1).Value
            ElseIf ws.Range("L" & r + 1).Value > Max_Total_Volume Then
                Max_Total_Volume = ws.Range("L" & r + 1).Value
                Max_Total_Volume_Ticker = ws.Range("I" & r + 1).Value
            End If
        Next r

 'Print greatest % increase, greatest % decrease, greatest total volume and their ticker names
        ws.Range("P2").Value = Max_Percent_Increase_Ticker
        ws.Range("P3").Value = Max_Percent_Decrease_Ticker
        ws.Range("P4").Value = Max_Total_Volume_Ticker
        ws.Range("Q2").Value = Max_Percent_Increase
        ws.Range("Q3").Value = Max_Percent_Decrease
        ws.Range("Q4").Value = Max_Total_Volume
        ws.Range("Q2:Q3").NumberFormat = "0.00%"

    Next ws
End Sub


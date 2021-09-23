'AUTHOR: Pablo Crespo Carrillo
'UNIVERSITY: The University of Western Australia
'COURSE: Data Analytics Boot Camp
'
'DESCRIPTION: 2nd assignment where I used VBA scripting to analyse real stock market data. The code loops
' through all the stocks for one year and outputs a new table with the following information (as described in the homework instructions):
'   - *Ticker Symbol*.
'   - *Yearly change* from opening price at the beginning of a given year to the closing price at the end of that year. A conditional 
'      formatting highlights positive change in green and negative change in red.
'   - *Percent change* from opening price at the beginning of a given year to the closing price at the end of that year.
'   - *Total stock volume* of the stock.
'
'As part of the instructions I also created a third table containing the stocks with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
'
'Finally, the script runs on every worksheet and displays a message box so the user knows the name of the active worksheet.

Sub stocks()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        Worksheets(ws.Name).Activate
        MsgBox ws.Name

    '   Defining last row dynamically -------------------------------------------------
        Dim Max_Row As Long
    '   Find the last non-blank cell in column A(1)
        Max_Row = Cells(Rows.Count, 1).End(xlUp).Row
    '   -------------------------------------------------------------------------------
        
        Dim counter, i As Long
        Dim opening_price, closing_price, greatest_increase, greatest_decrease, greatest_volume As Single
        Dim total_stoack As Long
            
'       Defining headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        counter = 2
        opening_price = Cells(2, 3).Value
        total_stock = 0
        
        For i = 2 To Max_Row
        
            total_stock = total_stock + Cells(i, 7).Value   'Calculating cumulative Total Stock volume for each Ticker
        
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then    'The conditional is True whenever the Ticker in row i is different from the Ticker in row i+1
            
                closing_price = Cells(i, 6).Value 
                
                Cells(counter, 9).Value = Cells(i, 1)   'Printing Ticker-value in Ticker-column
                Cells(counter, 10).Value = closing_price - opening_price    'Calculating and printing Yearly change in appropriate column

'               Conditional created to overcome division-by-0 problem since the conditional using function IsError() didn't work:
'                If IsError((closing_price - opening_price) * 100 / opening_price) Then ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                If opening_price = 0 Then
                    Cells(counter, 11).Value = Null
                Else
                    Cells(counter, 11).Value = Format((closing_price - opening_price) * 1# / opening_price, "Percent")     'Calculating Percent change
                End If
'               ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                Cells(counter, 12).Value = total_stock      'Printing Total Stock value
                
    '           Colouring conditionals ==================================================================
                If Cells(counter, 10).Value >= 0 Then
                     Cells(counter, 10).Interior.ColorIndex = 4 'Green (4) for positive
                Else
                    Cells(counter, 10).Interior.ColorIndex = 3  'Red (3) for negative
                End If
    '           =========================================================================================
    
                opening_price = Cells(i + 1, 3)         'Saving opening price of the next Ticker
                counter = counter + 1
                total_stock = 0
                
            End If
                
        Next i
        
    '   MsgBox (counter)
    
    '   CHALLENGE 1 *****************************************************************************************
        greatest_increase = Cells(2, 11)
        greatest_decrease = greatest_increase
        greatest_volume = Cells(2, 12)
        
        For i = 2 To counter - 1
    
            'Finding information of the Ticker with the Greatest % increase
            If Cells(i, 11).Value >= greatest_increase Then
                greatest_increase = Cells(i, 11)
                Range("P2").Value = Cells(i, 9).Value
                Range("Q2").Value = Format(greatest_increase, "Percent")
            End If

            'Finding information of the Ticker with the Greatest % decrease
            If Cells(i, 11).Value <= greatest_decrease Then
                greatest_decrease = Cells(i, 11).Value
                Range("P3").Value = Cells(i, 9).Value
                Range("Q3").Value = Format(greatest_decrease, "Percent")
            End If

            'Finding information of the Ticker with the Greatest total volume
            If Cells(i, 12).Value >= greatest_volume Then
                greatest_volume = Cells(i, 12).Value
                Range("P4").Value = Cells(i, 9).Value
                Range("Q4").Value = greatest_volume
            End If
    
        Next i
    '****************************************************************************************************************
    
    Next ws

    MsgBox ("End")
End Sub
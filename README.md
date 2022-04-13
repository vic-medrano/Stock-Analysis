# Stock-Analysis
VBA
## Overview of Project
This project was intended to help Steve analyze stocks to better help his parents with their financial decisions. He wanted to examine each stock by looking at their annual volume and return to make a more informed decision. 

## Results
When looking at the stocks in 2017 and 2018, we can see that the large majority of stocks in 2017 had positve returns. Four stocks (DQ, ENPH, FSLR, and SEDG) even had a return percentage that was greater than 100%. One of which, DQ, was the stock that Steve's parents were interested in. Only the stock, TERP, had a negative return. However, in 2018 all but two stocks did not perform well and had a negative return. The performance of the stocks for both years can be seen below to gain a better understanding. 


![VBA_Challenge_17](https://user-images.githubusercontent.com/102090016/163093016-173b013f-3d5c-4b88-a7c7-6160bcf4aa26.png)
![VBA_Challenge_18](https://user-images.githubusercontent.com/102090016/163093050-18ce2e81-0444-49c4-97f5-31dcd70e3772.png)
### Original Code

    
Sub AllStocksAnalysis()

Dim startTime As Single
    Dim endTime As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer

'1) Format the output sheet on the All Stocks Analysis worksheet.
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (2018)"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2) Initialize an array of all tickers
    Dim tickers(11) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
'3a) Initialize variables for the starting price and ending price
    Dim startingPrice As Single
    Dim endingPrice As Single
    
'3b) Activate the data worksheet
    Worksheets("2018").Activate
    
'3c) Find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row    
'4) Loop through the tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
    '5a) Loop through rows in the data
        Worksheets("2018").Activate
        For j = 2 To RowCount            
    '5b)Find total volume for the current ticker
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        End If
     '5c)Find starting price for the current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
        End If
        
    '5d) Find ending price for the current ticker
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
        End If
    Next j
        
'6) Output the data for the current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i
   
   endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


Sub formatAllStocksAnalysisTable()

'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#, ##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
   dataRowStart = 4
   dataRowEnd = 15
   For i = dataRowStart To dataRowEnd
   
    If Cells(i, 3) > 0 Then
    'Color the cell green
    Cells(i, 3).Interior.Color = vbGreen
    ElseIf Cells(4, 3) < 0 Then
    Cells(i, 3).Interior.Color = vbRed
    Else
    Cells(4, 3).Interior.Color = xlNone
    End If
Next i
    
End Sub

Sub formatAllStocksAnalysisTable()

'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#, ##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
   dataRowStart = 4
   dataRowEnd = 15
   For i = dataRowStart To dataRowEnd
   
    If Cells(i, 3) > 0 Then
    'Color the cell green
    Cells(i, 3).Interior.Color = vbGreen
    ElseIf Cells(4, 3) < 0 Then
    Cells(i, 3).Interior.Color = vbRed
    Else
    Cells(4, 3).Interior.Color = xlNone
    End If
Next i
    
End Sub

### Refactored Code
Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Count the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker index
    Dim tickerIndex As Integer
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
    tickerVolumes(tickerIndex) = 0
    
   
    Worksheets(yearValue).Activate
        
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
            '3b) Check if the current row is the first row with the current ticker.
                    
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
            
        '3c) Check if the current row is the last row with the current ticker.

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                'if the next row's ticker doesn't math, increase the tickerIndex
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        '3d) Increase tickerIndex
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
        Next i
        
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        

    
    Next i
    
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
                       
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub

Once the code was refactored the run times could be seen to take place at a significantly faster speed. The refactored run times had an almost 80% increase in speed. 


![Original_2017](https://user-images.githubusercontent.com/102090016/163187284-44788b1f-7fd4-4d07-9f34-0d16bee6bc9e.png)
The original run time for the 2017 analysis


![VBA_Challenge_2017](https://user-images.githubusercontent.com/102090016/163187843-8d76e576-89c0-4ed1-bfa4-3e46862ab4f5.png)
The refactored run time for the 2017 analysis


![Original_2018](https://user-images.githubusercontent.com/102090016/163187453-539a96d8-7b88-48b8-8d18-6dae544e95b4.png)
The original run time for the 2018 analysis

![VBA_Challenge_2018](https://user-images.githubusercontent.com/102090016/163187953-3c523f72-3a18-4ea0-92bc-9a1111ac931c.png)
The refactored run time for the 2018 analysis


## Summary
### Advantages/Disadvantages of Refactoring code
Advantages of refactoring code is that it allows the code to become more efficient and perform more quickly. This is useful for when analyzing mass amounts of data. The code is also easier to read which is ideal when many people are working on the same code.

Although there are many benefits, there are some disadvantages as well. Refactoring may make it more difficult to find an error as it is more complicated. The complications may also may make more room for errors and/or bugs.
### How do these pros/cons apply to refactoring the original script?
The pros can be seen in the run time that was performed. However, both run times were still relatively fast. For me, it was easier to see the cons, especially as a budding analyst. Refactoring did make it more difficult to find my errors especially when trying to compare to the original code to see where I may have gone wrong.

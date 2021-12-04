# Green Stock Analysis
## Overview of Project
My client Steve is happy with the workbook I prepared for him. He can analyze an entire dataset with a click of a button. He is wanting to do some more research for his parents.
He is wanting to expand the dataset exponentially to include the entire stock market over the last few years. The previous code I wrote him worked for a few stocks but it might not work well for thousands of stocks. We need to find a way to make this new workbook just as seamless as the first one I did for him. I will need to find a way to refactor my existing code to work through the dataset quickly but efficiently. To make sure I will get the results I am looking for, I will be comparing the time each workbook takes to finish.
# Results
## Refactoring the Code
I made changes to the original code to be more efficient. I created 3 new arrays: 
- tickerVolumes(12) to hold voulume
- tickerStartingPrices(12) - to hold the starting price
- tickerEndingPrices(12) - to hold the ending price

The tickers array I created in the original workbook established a ticker symbol that was called on in each stock. When I refactored the tickers array I was able to store performance data for each stock when a for loop runs an analysis on them.
The tickerIndex variable I created was able to match the 3 performance arrays so I can use nested for loops and variable to loop through the data and complete the analysis. This function allows me to get the results I want.

See the Refactored and Original coding below.

## Refactored Code

Sub AllStocksAnalysisRefactored()

     'Define starTime and endTime as variables
    Dim startTime As Single
    Dim endTime  As Single

    'Ask the client what year they would like to run with input box.
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
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        Dim tickerIndex As Single
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
    
        Dim tickerStartingPrices(12) As Single
   
        Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
            For i = 0 To 11
            tickerVolumes(i) = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
            For j = 2 To RowCount
    
    '3a) Increase volume for current ticker
11
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
    '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(j - 1, 1).Value <> Cells(j, 1) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                     
        'End If
            End If
        
    '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         'If  Then
            If Cells(j + 1).Value <> Cells(j, 1) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
    '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        'End If
        End If
       Next j
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
            For i = 0 To 11
        
            Worksheets("All Stocks Analysis").Activate
            
            tickerIndex = i
            
            Cells(4 + i, 1).Value = tickers(tickerIndex)
            Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
            Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
    
    
    'Format headline with bold and underline
            Worksheets("All Stock Analysis").Activate
            With Range("A3:C3")
            .Font.Bold = True
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            End With
        
        
    'Format number results on All Stock Analysis worksheet
            Range("B4:B15").NumberFormat = "#,##0"
            Range("C4:C15").NumberFormat = "0.0%"
            Range("B3:B15").AutoFit
            
        Next i
    'Add green and red background for positive and negative results, respectively. Use a For loop and conditionals.
    'Define variables for first and last row we are adding background color.
    
            dataRowStart = 4
            dataRowEnd = 15

    'Start a for loop
    
            For i = dataRowStart To dataRowEnd
            
            'Start conditional
        
                If Cells(i, 3) > 0 Then
            
                Cells(i, 3).Interior.Color = vbGreen
            
                Else
        
                Cells(i, 3).Interior.Color = vbRed
            
            'End conditional
            
                End If
                
    'End for loop
       
        Next i
 
 'End Timer and print msg on how long it took to run the code.
 
  endTime = Timer
  MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

## Original Code

Sub allstocksanalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

     startTime = Timer
     endTime = Timer
     MsgBox "This code ran in " & (endTime - starTime) & " seconds for the year " & (yearValue)
     
       
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    
    
   '2) Initialize array of all tickers
   
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
    
   '3a) Initialize variables for starting price and ending price
        Dim startingPrice As Single
        Dim endingPrice As Single
        
   '3b) Activate data worksheet
         Worksheets(yearValue).Activate
                
   
   '3c) Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
             
   '4) Loop through tickers
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
       
    '5) loop through rows in the data
            Worksheets("2018").Activate
            For j = 2 To RowCount
        
           '5a) Get total volume for current ticker
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
                
            '5b) get starting price for current ticker
            
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
                
            '5c) get ending price for current ticker
            
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                    
                End If
                
            Next j
            '6) Output data for current ticker
                Worksheets("All Stocks Analysis").Activate
                Cells(4 + i, 1).Value = ticker
                Cells(4 + i, 2).Value = totalVolume
                Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
     
        Next i

    
End Sub


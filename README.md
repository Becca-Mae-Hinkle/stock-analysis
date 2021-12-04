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
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    

    '1b) Create three output arrays
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
             tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
           If Cells(j - 1, 1).Value <> Cells(j, 1) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value  
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(j + 1).Value <> Cells(j, 1) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        'If  Then
                      

            '3d Increase the tickerIndex.
            
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
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


## Original Code

Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer

'1) Format the output sheet on All Stocks Analysis Worksheet

    'Activate "All Stocks Analysis" worksheet
    Worksheets("All Stocks Analysis").Activate

    'Title Analysis
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a Header Row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2)Initialize an array of all tickers.
    
    'Declare an array with 12 string elements
    Dim tickers(12) As String
    
        'Assign tickers to an element in the array
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
        
'3) Prepare for the analysis of all tickers.

    '3a) Initialize variables for the starting price and ending price.
    
        'Creating a Variable for Starting & Ending Price
        Dim startingPrice As Double
        Dim endingPrice As Double
    
    '3b) Activate the data worksheet.
        
        Worksheets(yearValue).Activate
        
    '3c) Find the number of rows to loop over.
        
        rowStart = 2
        'DELETE: rowEnd = 3013
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
'4) Loop through the tickers.
    
    For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0

'5) Loop through the rows in the data.

        'Activate Data Worksheet
        Worksheets(yearValue).Activate
        
        For j = rowStart To RowCount
        

    
    '5a) Find the total volume for the current ticker.
    
            'Identify ticker
            If Cells(j, 1).Value = ticker Then
                
                'increase ticker totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
    '5b) Find the starting price for the current ticker.
    
            'Identify first row of ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                'set starting price
                startingPrice = Cells(j, 6).Value
                
            End If
            
    '5c) find the ending price for the current ticker.
    
            'Identify last row of ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'set ending price
                endingPrice = Cells(j, 6).Value
                
            End If
            
        Next j
        
    
    
'6) Output the data for the current ticker.

        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker header
        Cells(i + 4, 1).Value = ticker
    
        'Sum of Volume
        Cells(i + 4, 2).Value = totalVolume
    
        'Return Value
        Cells(i + 4, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

   

End Sub

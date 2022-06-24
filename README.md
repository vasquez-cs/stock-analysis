# Refactoring with the Stock Analysis VBA code

## Overview of Project: 

For this challenge, we are looking to imporve upon the original file that was provided to Steve. To do this we refactored the code used in Module 2 solution code. This was done by looping through all the data one time, using code efficiently and making it easier to compherend while collecting the same information. Following the refactor the run times where analyzed to determine on whether the refactor was successful in making the VBA script faster. 

## Results
In our original data run within the green_stocks xlsm file we can see that the orignal AllStocksAnalysis sub macro for both the 2017 and 2018 runs are nearly one second as shown in both figure 1 and figure 2 below. While these appear to be fairly quick run times, we attempted to improve upon this by making use of a single for loop with multiple if else statements as shown in figure 3 below. This single for loop contains the main portion of the refactoring to inmprove speed. The single for loop incorporates several if statements that allow us to make use of the following arrays: tickerVolumes, Ticker, tickerStartingPrices, and tickerEndingPrices. The if statements set the ticker to 0 then moves on to identify starting and ending price, while maintaining the ticker number and finally outputing to the return column rows for appropriate tickers. As we can see in figure 4 and 5, our run times for the VBA refactored code was greatly improved for both 2017 and 2018. As shown in figure 6, there is almost an 87% decrease in the 2018 runtime! That is a great improvement that shows that our refactoring was successful.  

<img src="https://user-images.githubusercontent.com/107224632/175447188-9885726e-f925-4bfd-bfa9-a9c591b7dbee.png" width=40% height=40%><br />
*Figure 1: 2017 original data run time*

<img src="https://user-images.githubusercontent.com/107224632/175447443-2b841119-70c8-48f9-97aa-8b632a10060a.png" width=40% height=40%><br />
*Figure 2: 2018 original data run time*
        
        For i = 0 To 11
        
        tickerVolumes(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    'start of conditional, have rowcount start at "A1" by using cells() to avoid ticker header in A1.
    
        For i = 2 To RowCount
            
            If Cells(i, 1).Value = tickers(tickerIndex) Then
            
                '3a) Increase volume for current ticker
                
                'TickerIndex value goes up by adding the tickerVolumes value at current tickerindex with the existing value in the current "Volume" cell
                
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
                End If
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'Next conditional, starts by checking that the previous cell doesnt match the current ticker value and that the current cell does equal current ticker value. This ensure it is the first row.
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            'this pulls the value from the previous cell of the "close" and stores it in tickerStartingPrices since it would be the starting price for the first row.
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'Alter code for first row above to make it check for the last row. Instead of looking at the previous row, the plus 1 looks to the next row to ensure that there is not a matching ticker
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            'this tickerIndex starts at one and with each loop gets +1 until it hits the 11th index.
            
                    tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    'restate the worksheet I want to print out the Ticker, Total Daily Volume, and create a for loop to place the stored array values into the desire cell by adding +4, representing the header,  to i (represents the index number)
    
    Worksheets("All Stocks Analysis").Activate
    
        For i = 0 To 11
            
            Cells(i + 4, 1).Value = tickers(i)
        
            Cells(i + 4, 2).Value = tickerVolumes(i)
            
            Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

*Figure 3: Refactored VBA Code showing single for loop*

<img src="https://user-images.githubusercontent.com/107224632/175446694-f56472e5-ca3d-4111-9de1-a66b46ae9ca8.png" width=40% height=40%><br />
*Figure 4: 2017 Refactored data run time*

<img src="https://user-images.githubusercontent.com/107224632/175446779-0739a98a-86d2-47f7-9ec4-45db5dfe4da1.png" width=40% height=40%><br />
*Figure 5: 2018 Refactored data run time*

<img src="https://user-images.githubusercontent.com/107224632/175448831-63b07cfb-f29e-4ed7-9694-6ebb12730d4b.png" width=40% height=40%><br />
*Figure 6: Percentage decrease between orignal 2018 VBA code and 2018 Refactored VBA code run time* <sup>a</sup><br />



## Summary:
Refactoring code is beneficial when the updates made are efficent and produce a quicker runtime when using the same dataset for comparing the orignial code vs the new refactored code. In this case, the pros of refactoring the VBA code is that we produced nearly a 87% decrease in runtimes for the 2018 dataset. A potential con is when this refactored code is further altered to accomodate more tickers or using larger datasets. Further updates need to take into account that there may be ever more ways to make this efficient that do not use the current VBA code structure.

### Footnotes

<sup>a</sup> </sup>Screenshot from https://www.calculatorsoup.com/calculators/algebra/percentage-decrease-calculator.php </sup>

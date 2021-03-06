# Stock-Analysis
An analysis of stock data with Excel VBA

## Overview
### Purpose and Background
In this project, we worked with stock data provided by our client, Steve, in order to quickly compare the returns over a two year span. The aim of this analysis is to create a more efficient VBA script in Excel that has the potential to accommodate a vast amount of data from many more stocks.
 
In order to acheive this goal, we will edit, or refactor, the code that we initially came up with  during the module. The goal is to create a script that, upon user input, can loop throught the entirety of the data and print the correct values into an analysis sheet. Though only testing with a dozen stocks, this created script should have the potential to accommodate any number of stocks as well as a different year values (provided the user has access to the data) and run faster - more efficiently - than the previous script.

## Results
In the images below, we see that the stock returns for 2017 were much better compared to 2018. In the latter year, all but two of the stocks had negative returns, meaning investor most likely lost money that year. However in 2017, only one stock came out with negative returns, and a third of the others had over 100% in returns. 

The images also depict the runtime for the script both before and after being refactored. We can see that before refactoring the code, the runtime was nearly 1 second, Only decreasing to the values shown below after multiple runs in quick succession. However, after refactoring, the run time came to be at about a third of the original for both years when tested. This indicates that the changes made to the code were successfull in acheiving one of our goals.

| Runtime for 2017 data - _before_ refactoring | Runtime for 2018 data - _before_ refactoring |
| ------------- | ------------- |
| ![](https://github.com/chichi-ugo/stock-analysis/blob/main/Resources/VBA_Challenge_2017_module.PNG?raw=true) | ![](https://github.com/chichi-ugo/stock-analysis/blob/main/Resources/VBA_Challenge_2018_module.PNG?raw=true) |
| ------------- | ------------- |
| **Runtime for 2017 data - _after_ refactoring** | **Runtime for 2018 data - _after_ refactoring** |
| ![](https://github.com/chichi-ugo/stock-analysis/blob/main/Resources/VBA_Challenge_2017_refactored.PNG?raw=true) | ![](https://github.com/chichi-ugo/stock-analysis/blob/main/Resources/VBA_Challenge_2018_refactored.PNG?raw=true) |

We can see that having the program assign our output variables into an array helps to speed up the overall process of the code. This is most likely because the program is now able to run through all the data and then print out the outputs at the end, rather than runing through the data and each time printing out a data point before moving to the next. A similiar analogy is: when at the grocery store, you put all your items in a basket and then bring it up to the register at the end instead of grabbing one item and then taking it up and going back for the next item and so on.

```
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
```   
Above shows how we define the arrays. These lines are put before we ask the program to run through the data to collect the information. 

```
  '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
```
After the program runs throught the data in the sheets and has stored its outputs, we simply ask the program to print its outcomes with a separate `for` loop - as shown above.

When we look at our original - unrefactored code - we see that the outputs are just assigned to a variable. This means the program will have to print that variable before reassigning it for the next run in our loop.

```
'Loop through the tickers
    For i = 0 To 11
        ticker = tickers(i)
        totVol = 0

        'Loop through rows in the data
        'Worksheets("2018").Activate - commented this line out when testing yearValue
        Worksheets(yearValue).Activate
        
        For j = 2 To RowCount

            'Find the total volume for the current ticker
            If Cells(j, 1).Value = ticker Then
                totVol = totVol + Cells(j, 8).Value
            End If

            'Find the starting price for the current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startPrice = Cells(j, 6).Value
            End If

            'Find the ending price for the current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endPrice = Cells(j, 6).Value
            End If

        Next j

    'Output the data for the current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totVol
    Cells(4 + i, 3).Value = endPrice / startPrice - 1

    Next i
```

## Summary
1. What are the advantages or disadvantages of refactoring code?
    - Advantages: Refractoring the code not only helps to make the program run faster in some instances, it also helps to pare down the code and make it more straightforward and easier to interpret. This is important because in the instance that the code is further adapted, it will be easier to debug and more adaptable to changes. Similarly, it helps other users to understand the processes taking place in your code and can help them understand their analysis better. 
      
    - Disadvantages: For certain long processes, it may not be as advantageous to try and put a lot of information/tasks into one routine. For some, it may be clearer to have the tasks broken up into the different routines in order to retain clarity. Furthermore, when refactoring to accommodate broader testing ranges, you may not be equipped with enough data to properly test the code and ensure that there are no bugs.
 
2. How do these pros and cons apply to refactoring the original VBA script?
    - In our analysis, we saw a significant decrease in overall runtime of the program with the refactored code, which is a major advantage over the original. We also can note that the final refractored code is a bit more concise than the original, as it contains, overall, less lines of code. This could be beneficial to other who may want to adapt the code in the future. A con of the refactored code could be that although we managed to take out many of the hard coded values, there were still some left. Though this does not have a negative effect on our program or outputs in this study, it does present a slight obsticle when considering handling a dataset with many more stocks to account for. 

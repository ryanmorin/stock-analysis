# Refactor VBA code and Measure Performance

## Overview of Project
The project uses daily stock data to calculate the total volume and performance for 12 companies. The stocks are identified by their 4 character ticker symbol. The script also performs some formating identifying those stocks that had positive performance in green, and negative performance in red. The goal of the project is to make the [AllStocksAnalysis](https://github.com/ryanmorin/stock-analysis/blob/main/AllStockAnalysis) code developed in Module 2 run faster by making some changes.

### Purpose
Reducing the code run time will make the program more useful because it will run faster on larger data sets without reducing the output. This can be accomplished by eliminating one of the 'For' loops and replacing it with arrays.

## Analysis

### Analysis of Problem and Strategy to Reduce the Script Run Time
The existing AllStockAnalysis code uses two 'For' loops: one loop to run through 12 ticker symbols and a second loop to loop through more than 3000 rows of daily stock data.  Running a nested 'For' loop multiplies the number of rows that VBA processes. In total, the VBA script will run through 36,000 rows (12 * 3000).

The portion of the AllStocksAnalysis code with the nested for loops:
```
For j = 0 To 11

   ticker = tickers(j)
   totalVolume = 0
   
   Worksheets(yearValue).Activate
   For i = rowStart To rowEnd
```

We can reduce the number of 'For' loops used and thereby the amonut of time needed to run the script by replacing the 'For' loop with arrays. Once the remaining loop is finished running, the stock calculations can be performed on the data stored in the arrays. This change will reduce the number of rows that VBA needs to process from 36,000 to 3012.

The same section of code as above deleting the nested for loop and adding 3 new arrays:

```
'1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For init = 0 To 11
        tickerVolumes(init) = 0
        tickerStartingPrices(init) = 0
        tickerEndingPrices(init) = 0
    Next init

    ''2b) Loop over all the rows in the spreadsheet.
    For m = Start To RowCount
```

## Results

Deleting the for loop and replacing it with arrays, reduced the run time for the 2017 and 2018 data.

### Summary of the Run Time 2017

**The Original AllStocksAnalysis Code Run Time 2017**

![2017_original_run_time](https://github.com/ryanmorin/stock-analysis/blob/main/original_2017_run_time.png)

**The Refactored AllStocksAnalysis Code Run Time 2017**

![2017_refactored_run_time](https://github.com/ryanmorin/stock-analysis/blob/main/refactored_2017_run_time.png)

After the changes to the code the refactored code ran 75% faster compared to the original. The output was unchanged.  See the side by side comparison below.

![2017_comparison](https://github.com/ryanmorin/stock-analysis/blob/main/2017_comparison.png)

### Summary of the Run Time 2018

**The Original AllStocksAnalysis Code Run Time 2018**

![2018_original_run_time](https://github.com/ryanmorin/stock-analysis/blob/main/original_2018_run_time.png)

**The Refactored AllStocksAnalysis Code Run Time 2018**

![2018_refactored_run_time](https://github.com/ryanmorin/stock-analysis/blob/main/refactored_2018_run_time.png)

After the changes to the code the refactored code ran 76% faster compared to the original. The output was unchanged.  See the side by side comparison below.

![2018_comparison](https://github.com/ryanmorin/stock-analysis/blob/main/2018_comparison.png)

## Summary

There are a number of advantages associated with refactoring code; I'll focus on two.  Our example highlights how refactoring code often improves the performance of the code by making it run faster.  This frees up tech infrastructure to focus on other tasks thereby improving overall efficiency and maybe even cost.  One of the other advantages associated with refactoring code is it might allow a different set of eyes to look at a problem. Returning to a problem or having someone else refactor code might uncover bugs that weren't originally detected. The benefit is that this may provide more accurate results.

The disadvantages associated with refactoring code is the time and money associated with having a person or a group of people revisit something that's already complete. There's no guarantee that the code will work better. Further, any cost savings associated with improved code will be offset by the cost associated with the effort to refactor the code in the first place.  Finally there might not be enough time to dedicate to refactoring code. Programming resources are frequently stretched thin, so working on something new might be more important to the company or the programmer than revisitiing the past.

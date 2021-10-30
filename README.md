# Refactor VBA code and Measure Performance

## Overview of Project
The project uses daily stock data to calculate the total volume and performance for 12 companies. The stocks are identified by their 4 character ticker symbol. The script also performs some formating identifying those stocks that had positive performance in green, and negative performance in red. The goal of the project is to make the [AllStockAnalysis](https://github.com/ryanmorin/stock-analysis/blob/main/AllStockAnalysis) code developed in Module 2 run faster by making some changes.

### Purpose
Reducing the code run time will make the program more useful because it will run faster on larger data sets without reducing the output. This can be accomplished by eliminating one of the 'For' loops and replacing it with arrays.

## Analysis

### Analysis of Problem and Strategy to Reduce the Script Run Time
The existing AllStockAnalysis code uses two 'For' loops: one loop to run through 12 ticker symbols and a second loop to loop through more than 3000 rows of daily stock data.  Running a nested 'For' loop multiplies the number of rows that VBA processes. In total, the VBA script will run through 36,000 rows (12 * 3000).  

We can reduce the number of 'For' loops used and thereby the amonut of time needed to run the script by replacing the 'For' loop with arrays. Once the remaining loop is finished running, the stock calculations can be performed on the data stored in the arrays. This change will reduce the number of rows that VBA needs to process from 36,000 to 3012. 

## Results

Deleting the for loop and replacing it with arrays, reduced the run time for the 2017 and 2018 data.



Next, the analysis focued on the goals and how that impacted the outcome. The smaller the goal, the greater the chance for the campaign to meet it's goal. Successful campaigns were on average 6x's smaller than unsuccessful campaigns. In the Theater category, the average successful campaign was over 30% smaller than the average and the average unsuccessful campaign was 9% larger. Because the average successful Theater goal was smaller compared to the average, Theater required fewer backers to meet it goal. 

There are two fields that require additional research: Staff Pick and Spotlight. Both of these fields appear to be closely associated with a campaign's success. In the case of Spotlight, if that field was coded as 'True', 100% of the campaigns were successful. Understanding how the Spotlight or Staff Pick field is assigned would be very useful in recommending a success strategy for campaigns. 
   
There are some other graphs that would be useful:
   1. Average amount of time for a successful campaign
   2. Average amount of goal for successful campaigns vs. unsuccessful campaigns




[Kickstarter Challenge](https://github.com/ryanmorin/kickstarter_analysis/blob/main/kickstarter_challenge.zip)

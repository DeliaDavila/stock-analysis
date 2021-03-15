# stock-analysis
Using VBA to analyze stock data in Excel

## Overview of Project
### Purpose
The purpose of this project was to learn how to refactor code as needs arrive. In this case, the goal was to reduce the runtime of the code. Additionally, I learned a method for achieving a similar result without using a nested loop.
### Background
We had previously completed a stock analysis subroutine that looped through an array of stock tickers and created a table showing the total daily volume and calculated the yearly return from a requested input. On my computer, this code is taking about .6 seconds to run. If the refactoring is successful, the runtime should be decreased.
## Process
I used the previous code to begin working on a new routine, *AllStocksAnalysisRefactored*. Initially, I kept the nested loop from the previous code. The code resulted in the correct outcome but at essentially the same speed. The hints suggested a better method and in the office hours, a teaching assistant explained arrays visually which helped me understand how the code was populating the arrays.


### Loop 1, set initial volumes to zero
The first loop goes through the array *tickerVolumes* and sets the value to zero. This sets the initial state so that the calculations can begin without having to set and reset the volume for each new ticker.
‘‘‘
    For h = 0 To 11
        tickerVolumes(tickerIndex) = 0
    Next h
‘‘‘
### Loop 2, populate the arrays with calculated values
The second loop goes through the rows for the initial ticker calculating the total volume by adding to *tickerVolumes* for each line of data for the ticker. To calculate the return, the code finds the first and last row with data for the ticker and calculates the difference between starting and ending values. This is largely kept the same from the previous code, except for substituting arrays for individual variables. One significant change is an addition to the instructions when the last row of data is found. In addition to using the data to calculate the return, the tickerIndex is increased by one (see code below). This happens for “Next i” starts the loop again, so the code will now be populating the data for the next array. Because we set the initial condition with the first loop, the volume calculation for the next ticker starts over at zero.
‘‘‘
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    Next i
‘‘‘

### Loop 3, pull the calculated values from the array into the spreadsheet
Once the arrays have been populated with information from the second loop, the third loop puts those values into the spreadsheet. To do this, I created a loop with the number of items in the array and used a loop to go through them. The code is similar to what was used previously but now uses the loop to both pull from the correct place in the array and populate different rows in the spreadsheet, with the correct ticker pulled from the tickers array.
‘‘‘
    For i = 0 To 11
    'output the information in the arrays
        Worksheets("All Stocks Analysis").Activate
        Cells((i + 4), 1).Value = tickers(i)
        Cells((i + 4), 2).Value = tickerVolumes(i)
        Cells((i + 4), 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
    Next i
‘‘‘

## Results
The refactored code successfully reduces the runtime of the code. This is easily demonstrated by using the buttons to run the previous code (Run All Stocks Analysis) and the new code (Run refactored). Where the previous code runs at about .6 seconds, the new code runs at around .1 second. See screenshots below.
![Screenshot2017](add link)
![Screenshot2018](add link)

## Summary
In this section, I discuss the relative merits of refactoring
### Advantages and Disadvantages of Refactoring Code in General
The general reason code is refactored would be to improve it or use it again. Improvements could be computer resources, reducing the time for the code to run, or to human resources, making the code clearer so that future use of the code is easier. This might be done to simply adjust for differing inputs, as we did with our first version of code where a specific year was hard coded and was switched to user input. This could also be done to simplify the code for future use. The use of nested loops could be very confusing in very complex code, so breaking down the needs into different sections would make it easier to use just part of the code in later projects. 
The disadvantages in refactoring are that it takes time to make adjustments in the script. If the script functions either way, then it may seem unnecessary to make the adjustments, especially since it’s not known whether the performance improvement will be significant. Additionally, changing code always introduces the chance for mistakes so refactored code will still have to be QA’d as if it were new code.
### Advantages and Disadvantages of Refactoring of this script
This particular code was refactored from original to improve the speed.  The refactored VBA code runs at about one-sixth the speed of the original. An additional advantage is that the actions of the script have been separated into different processes. This might make re-using the code easier in the future.
As for the disadvantages for this script, since this was a training exercise, the additional time spent on refactoring had its own goal of learning, so that time was necessary. The time spent debugging each other’s work during the office hours indicated that many small mistakes were made during the refactoring. But again, while this would be a disadvantage in the real world, it was an advantage here as it gave us a chance to learn what some errors mean and learn from the teaching assistants how to work through the problems, debug, and root out minor issues like typos.


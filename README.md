# Stock Analysis

[VBA_Challenge](/VBA_Challenge.xlsm)

## Overview of Project
Steve would like to do more research on stocks for his parents.  I will use VBA to analyze a portfolio of stocks and their performances in 2017 and 2018.  The code is already written, but I have refactored it to run more efficiently for Steve.

## Results

### Writting the Code
The VBA code loops through all of the pertinent tickers in the index after initializing the array, which looks like this:

```
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
```

After setting the tickerIndex to zero and initializing the stock volumes to zero, I loop through row two and the final row on the desired Excel tab.  The middle of the "for loop" looks like this:

```
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```

It uses the index to increase the volumes for each ticker.  Deeper in the code, I increase the tickerIndex to move on to the next ticker only after successfully storing the starting and ending prices for the current ticker.
```
tickerIndex = tickerIndex + 1
```

The loop continues until we have what we need for all tickers.  Using the tickerIndex in this refactored code is much more efficient than my original version.  The analysis on 2017 and 2018 ran in 0.289 and 0.273 seconds respectively while the non-refactored versions ran in over 0.82 seconds for both years.

### Refactored 2017
![VBA Challenge 2017](/resources/VBA_Challenge_2017.png "VBA Challenge 2017")
### Refactored 2018
![VBA Challenge 2018](/resources/VBA_Challenge_2018.png "VBA Challenge 2018")
### Original 2017
![Original 2017](/resources/green_stocks_2017_Not_Refactored.png "Original 2017")
### Original 2018
![Original 2018](/resources/green_stocks_2018_Not_Refactored.png "Original 2018") 

### Analyzing Stock Performance
With the exception of TERP, most of the stocks in this portfolio performed really well in 2017.  However, in 2018, they all flipped to negative returns except for ENPH and RUN.  This is a mixed bag of results.  For further analysis, I would recommend diving into financial statements for ENPH and RUN to make sure that their strong two year performance is not a fluke.

## Summary

This refactored code makes use of indexes and arrays, which has plenty of advantages.  Not only is it less code and easier on the eyes, but it's much nicer to read if you hand it off to someone else to view.  Also, it runs much faster so it is all-around more efficient.  
  
The original VBA script might have been less difficult to understand at first.  It is more logical so every piece of it is written out as our brains work the problem.  However, writing every single step is clearly not as beneficial, evident by the run times.  When we make use of arrays and an index, we have to think about where to make the edits to strengthen the code, but it leads to a more sophisticated version of the script that is faster in the end.






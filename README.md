# Stock Analysis Using VBA - Refactored Code
## Overview and Purpose of the Project
Having learned the basics of VBA and how it can be applied to the comb through thousands of lines of data with the press of a button, it became apparent that there is a better, more efficient way to perform the same task. The method of streamlining code is known as "refactoring code", which is what was performed in this analysis. 

The original and refactored macros scan through the stock data tabulated for years 2017 and 2018, both of which contain over 3,000 rows of data. With the click of a "run" button, the macro would determine the starting and ending price for each of the 12 stocks for the year and subsequently calculate the return. It would then tabulate the total volume traded. All this information would then be printed on a separate sheet in a comprehensive format. The goal is for the refactored code to produce the same results which were performed in the sub titled, "AllStocksAnalysis()", located in the Module 1 macro of file, "VBA_Challenge.xlsm". In the same file, Module 3 holds our refactored code.
## Results
To be sure the measurements were fair, both macros incorporated a timer function that would begin after the desired year of analysis was determined, which would then stop and provide a message box with the results at the very end of the macro.
### Original Macro (Module 1 - Sub "AllStocksAnalysis()")
![All_Stocks_2017](https://user-images.githubusercontent.com/92493572/140650660-cec8bf3c-05c4-46ab-adb8-8bd77a1f404a.png)
![All_Stocks_2018](https://user-images.githubusercontent.com/92493572/140650665-2f5b053d-3536-44de-aaee-aec3f48688aa.png)

As seen above, the execution time appears very quick, but let's compare that to the refactored version.
### Refactored Macro (Module 3 - Sub "AllStocksAnalysisRefactored()")
![VBA_Challenge_2017](https://user-images.githubusercontent.com/92493572/140650847-2fcd4407-ac51-4be2-8be6-5b667c4328b6.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/92493572/140650848-6dffbf49-7d00-48db-b3e6-29e9da4af943.png)

The refactored code performed the same task much faster. In fact, the refactored code was 7.8 times faster at running the 2017 data, while it was exactly 8 times faster at running the 2018 data. This was determined by dividing the original run time by the refactored run time for the respective years.
### Causes of the Results
The way the original code and the refactored code loop through the 3,000 rows is where the change in speed comes from. The original code will scan all 3,000 rows looking to tabulate information for one stock. After it completes that task, it changes to the next stock and cycles through all 3,000 rows *again*. As there are 12 stocks (fortunately only 12) in the dataset, it loops through 3,000 rows 12 times. This proves to be a lot of extra work that the computer needs to perform. Before it cycles to the next loop it prints the results on to a separate sheet.

The refactored code loops through all 3,000 rows, but only needs to do it one time. The code recognizes where a new stock begins and stores all current stock information in its own *separate* array value before moving on to the next. When the entire loop is complete, it then prints all the saved array values to a separate sheet. When looping through the stocks, the original code is likely working 12 times harder.
## Summary
While the produced results cut the execution time to a small fraction, refactoring code should not necessarily be performed on all code. Multiple aspects must be considered when deciding if refactoring is the correct decision to make, or at least appears worthwhile.
  - There is no way to tell how long the refactoring will take (think 'opportunity cost').
  - How much is to be invested in the task before it becomes a financial burden.
  - Does this section of code *need* to be refactored? (for instance, is it only used once a year to perform simple tasks?)
  - Will the result prove to be more efficient?

Determining if refactoring code is worthwhile is partially subjective, though in a business environment it may be worthwhile to perform financial analysis including forecasts to weigh the options. Understanding that neat, tidy, and efficient code is what all code *should* be, sometimes there simply isn't a strong enough demand or resources to effect that change.

In this case, the refactored VBA script was able to consistently process the results around 8 times faster. If this code is of importance to the user, this is a substantial change. The refactored code is neater and no longer uses a nested 'for loop', which may otherwise more easily confuse someone reading the code. One notable disadvantage was the time it took to produce effective code. The first attempts actually ran slower than the original (roughly 0.85 seconds). In total, roughly 10 hours were spent producing and debugging it until the final result was produced. Depending on the scale of how the refactored code will be used, it may have absolutely been worthwhile, even considering the time investment needed to make it function.

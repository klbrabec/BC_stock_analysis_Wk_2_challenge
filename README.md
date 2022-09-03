# VBA of Wall Street - Challenge Week Two 

## Overview of Project


### Purpose
There are two requirements for this project.   The first is to validate that the code developed for the initial dataset analysis
works as expected for a larger data set, and provides the same information.  The second requirement is to ensure that the code 
performs well, and returns results as fast, or faster than it did prior to being refactored.    This is not new code development, 
it relies on work done previously that has been verified.  However, this code has been reviewed and refactored in order to be 
more efficient.  

## Results

### Refactoring Approach
The intial code was designed and constructed to provide the total daily volume of sales of a particular stock (ticker) as well as 
a percentage indicator of the return on the investment in that stock.  This was done by using logic that used a single For loop to 
access information regarding individual tickers.  Because of this design, the code would loop through a ticker, and then within that 
loop, would process through each row in the data set, increasing the total volume by the daily volume, and doing a comparision of the 
row before and after to determine if it was the first or last row in the set to set the starting price and ending price values.  

![VBA_Challenge_CodeSnippet1.png](https://github.com/klbrabec/stock_analysis/blob/main/VBA_Challenge_CodeSnippet1.PNG)

While this works, it is not efficient for larger data sets.  The process of refactoring kept the logic in place, but implemented 
arrays, which will allow for accessing data related to a specific ticker, rather than having to loop through multiple values to find 
values needed.  This limits the processing overhead and improves the performance of the code itself. 

![VBA_Challenge_CodeSnippet2.png](https://github.com/klbrabec/stock_analysis/blob/main/VBA_Challenge_CodeSnippet2.PNG) 

 
### Results  

The results returned for the pre-refactoring and post-refactoring of the code are the same.  The difference in the new code can be seen
in comparision of the performance or completion of the macro  after the refactoring compared to the results prior to the refactoring.  
for the 2017 data set, the results can be seen here: 

![VBA_Challenge_2017_BeforeRefactor.PNG](https://github.com/klbrabec/stock_analysis/blob/main/VBA_Challenge_2017_BeforeRefactor.PNG)
As you can see in this image, the code completed on the 2017 data set within 4.96 seconds prior to being refactored. 
![VBA_Challenge_2017_AfterRefactor.PNG](https://github.com/klbrabec/stock_analysis/blob/main/VBA_Challenge_2017_AfterRefactor.PNG)

After the code was refactored, the code completed on the 2017 data set within 0.15625 seconds. This is an approximately 96.8% reduction in 
cycle time for the same data set.  Note, the completion of a macro/sub routine can be impacted by a number of different factors, however the magnitude of the improvement should be fairly similar
 
The results for the 2018 data set are similar. 

![VBA_Challenge_2018_BeforeRefactor.PNG](https://github.com/klbrabec/stock_analysis/blob/main/VBA_Challenge_2018_BeforeRefactor.PNG)
Results for 2018 were returned in 4.207031 seconds.  
![VBA_Challenge_2018_AfterRefactor.PNG](https://github.com/klbrabec/stock_analysis/blob/main/VBA_Challenge_2018_AfterRefactor.PNG)
After refactoring, those same results were returned in  0.1328125 seconds.  Again, this was a 96.8% reduction in the time 
it took to process and return results.  

Being able to enhance the performance of a macro will allow for it to be used for larger data sets in the future. 

### Data Results

As shown in the images above, the code performed the same tasks as they processed through the For loops.  The results for 
the original and refactored code are the same.  

## Summary
Refactoring of code is a common activity in software development.  Developers refer to pre-existing code in order to deliver 
faster, and work more efficiently. Additionally, refactoring of code will allow for the re-use of complex logic rather than 
having to 'reinvent the wheel'.  It is important however, to ensure that any code that is being refactored is fully understood 
prior to being modified, especially if that code is being called in another location that may 'break' due to unforseen dependencies
on the refactored code.  These are generally considered regression issues. 

In both the original and refactored code, while they both do the same thing, they both have a similar shortcoming.  In order to add any 
new stock tickers or modifying existing stock tickers that are being analyzed, the code itself would need to be modified to implement the 
changes in the index (either adding a new index value or modifying an existing index value).  In order to make this code 
more flexible and useful as stock portfolios and the stock market itself grows and changes, the array 
that defines the various tickers should be designed to be flexible where possible, rather than containing a set number of static values.
There are a number of ways to do this, however some of the more popular methods include 
 - Additional Macro functionality such as - VBA ArrayList (https://excelmacromastery.com/vba-arraylist/), 
 - Reading values into an array from a single column (https://vbaf1.com/tutorial/arrays/read-values-from-range-to-an-array/) and
https://stackoverflow.com/questions/11504024/fastest-way-to-read-a-column-of-numbers-into-an-array) 

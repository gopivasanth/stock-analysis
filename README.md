# stock-analysis
---
## Overview of Project
Steve intends to help his parents with their investments. Hence we created a workbook for him to help with his research, however, he wants to expand the dataset across multiple years. Our original workbook may work good for smaller datasets but if the dataset is expanded it may take a long time to execute. 

### Purpose
The purpose of this analysis is to refactor the original VBA script in order to make it run more efficiently and decrease its execution time. 

## Analysis and Refactoring

The file and scripts; original and refactored can be found in the file below:

[VBA_Challenge.xlsm](https://github.com/gopivasanth/stock-analysis/blob/5fe13cc1464fff61632bf07b003ae0ea32a2c10a/VBA_Challenge.xlsm)

### Differences in the code

- The  original code looped through an array representing each of the stocks to be analyzed, within that it looped through all of the rows in the worksheet and found the information for the correct variables then displayed it. 
- The variables we were looking for were the total volume of stocks, the starting price for that year, and the ending price for that year.

After Refactoring;
- The refactored code did not have nested loops instead it had 3 separate loops, additionally, as opposed to one array there were stocks, volume of stocks, starting price, and ending price. 
- The first loop initialized the volume of stocks to zero each time. 
- The next one looped through each row and calculated the volume, start and end price. 
- The final loop output the calculated variables for each stock and recorded it in the specified worksheet. 
- Within the second loop an index that was externally set to zero was used to keep reference of which stock the loop iteration was looking for; this index was then increase after each loop iteration.

These changes in the refactored allowed the analysis to run more efficiently and there by decreasing the execution time.

## Results

### Analysis of 2017 Stocks: Original vs Refactored

The original code results in a slower execution time as seen in the image below.

![Original 2017 run time](https://github.com/gopivasanth/stock-analysis/blob/5fe13cc1464fff61632bf07b003ae0ea32a2c10a/Resources/2017_old.png)

The refactored code results in a faster execution time as seen in the image below.

![VBA_Challenge_2017](https://github.com/gopivasanth/stock-analysis/blob/5fe13cc1464fff61632bf07b003ae0ea32a2c10a/Resources/VBA_Challenge_2017.png)

### Analysis of 2018 Stocks: Original vs Refactored

The original code results in a slower execution time as seen in the image below.

![Original 2018 run time](https://github.com/gopivasanth/stock-analysis/blob/5fe13cc1464fff61632bf07b003ae0ea32a2c10a/Resources/2018_old.png)

The refactored code results in a faster execution time as seen in the image below.

![VBA_Challenge_2018](https://github.com/gopivasanth/stock-analysis/blob/5fe13cc1464fff61632bf07b003ae0ea32a2c10a/Resources/VBA_Challenge_2018.png)

### Comparing 2017 to 2018 Stocks

By looking at the images above we compare the return for each stock in each year and how they differed. 
- We can see that in 2017 most of the stocks had a positive return. 
- In fact, all stocks except TERP had a positive return with stocks DQ, ENPH, FSLR, and SEDG all having more than a 100% return. 
- This was NOT true for the year 2018. In 2018 only two stocks had a positive return: ENPH and RUN. All other stocks had a negative return. 

From this we can draw out the conclusion 
- ENPH stock is worth investing in since it had high positive returns for two years in a row: 129.5% in 2017 and 81.9% in 2018. 
- RUN stock may also be a good investment however, the magnitude of this positive return may not be as certain as the return was only 5.5% in 2017 but increased to 84% in 2018. 

## Summary

### Advantages and Disadvantages of Refactoring Code

The advantage of refactoring code is 
- cleaner and more efficient code.
- more readable and reusable. If another person wants to reuse the code, they should not find it too hard to read and understand as it is intended to streamline processes. 
- run more efficiently as well-meaning larger datasets should not take as long to analyze.

One disadvantage of refactoring code is that is can be time consuming, there is no way to know how long this process may take and, in some cases, you won't have the opportunity to take the time necessary to refactor the code.

### Pros and Cons of Original vs Refactored Script

The advantage of the original script is that 
- there are less arrays and variables to create and keep track of.
- this makes the overall code easier to keep track of. 

The disadvantage to this approach is that once the dataset larger this is no longer an advantage, the code becomes more complex.
- since each variable is set and there is more reliance on set values, as opposed to the code finding these variables for itself and resetting, with larger data sets the code will take longer to execute and may become quite confusing. 

The advantage of the refactored script is that
- it makes the analysis smoother. 
- it takes less time to execute, and it can be used for larger datasets where the arrays can be modified or expanded for as many different variables as necessary. 

This advantage is evident when comparing the execution times between the original vs refactored code shown in the images above.

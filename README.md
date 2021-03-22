# stock-analysis
Module 2 stock analysis

## Analyzing yearly volume and return for 12 stocks

### Overview of Project
The objective was to help Steve demonstrate to his parents a case to invest in 12 different stocks based on total yearly volume and return by the push of a button using VBA. This was ultimately a proof of concept for his goal to analyze the entire market over the last few years. 

## Analysis and Challenges
The analysis was performed by using financial data which included the ticker, total daily volume, and daily price information. I used the starting and ending prices to calculate the return for the year. One major challenge was that if the data was not in order by date, the outcomes would be off for the return amounts. The code can be updated to find the earliest and latest date per a ticker to avoid a mistake in calculations. 

The Excel file that shows the analysis is located here: [GitHub Pages]( https://github.com/trallen09/stock-analysis/blob/main/VBA_Challenge.xlsm)

### Analysis of Outcomes Based on 2017 & 2018 stocks using VBA
If only looking at 2017 as an indicator moving forward, a lot of money could have been lost in 2018. In 2017, "DQ" almost had a 200% return. Overall, 2018 was a bad year for all but two of the selected stocks ("ENPH" and "RUN"). Both stocks seemed to withstand the downturn in 2018 while "RUN" increased performance significantly.

The VBA code was visibly different in the execution from the original to the refactored. I could watch the fields populate in the original code as opposed to the refactored.

- What are the advantages or disadvantages of refactoring code?
  - The advantage is faster execution times. 
  - The major disadvantage is that refactoring requires an indepth knowledge of the language to understand how things can be simplified.
- How do these pros and cons apply to refactoring the original VBA script?
  - As a pro and seen below, the refactored code performed much faster than the original.
  - As a con, there was quite a bit of struggle understanding the use of indexing the outcomes. I can remove a portion of the code and it still works fine.


![alt text](https://github.com/trallen09/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![alt text](https://github.com/trallen09/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Results

- 2018 original run time:
  - .671875 seconds
- 2018 refactored run time:
  - .1171875 seconds
- 2017 original run time:
  - .6796875 seconds
- 2017 refactored run time:
  - .109375 seconds


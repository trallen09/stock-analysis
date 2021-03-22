# stock-analysis
Module 2 stock analysis

## Analyzing yearly volume and return for 12 stocks

### Overview of Project
The objective was to help Steve demonstrate to his parents a case to invest in 12 different stocks based on total yearly volume and return by the push of a button using VBA. This was ultimately a proof of concept for his goal to analyze the entire market over the last few years. 

## Analysis and Challenges
The analysis was performed by using financial data which included the ticker, total daily volume, and daily price information. I used the starting and ending prices to calculate the return for the year. One major challenge was that if the data was not in order by date, the outcomes would be off for the return amounts. 

The Excel file that shows the analysis is located here: [GitHub Pages]( https://github.com/trallen09/stock-analysis/blob/main/VBA_Challenge.xlsm)
### Analysis of Outcomes Based on 2017
Using parent category, outcomes, and date created, I was able to determine the best launch date specific to theater campaigns. Filtering on the parent category "theater", I summarized a pivot table based on launch date for rows and outcomes for both columns and count of outcomes.  

![alt text](https://github.com/trallen09/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
### Analysis of Outcomes Based on 2018
The use of COUNTIFS allowed me to set the criteria for plays, goal limits, and outcomes. I had to create different limits to capture outcomes based on goals that fell between 1000 to 50000 within the COUNTIFS function. All the data was filtered for plays.

![alt text](https://github.com/trallen09/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Results

- What are two conclusions you can draw about the Outcomes based on Launch Date?
  - Campaigns have the best chance to succeed in May. Failures and canceled campaigns are consistent throughout the year with October being the highest failures.
- What can you conclude about the Outcomes based on Goals?
  - The most successful campaigns are under $10,000 while most campaigns over $45,000 tend to fail. Over 80% of the campaigns are under $10,000 as well.
- What are some limitations of this dataset?
  - There could be errors within the data since I did not search for duplicates and other cleaning proceedures.
- What are some other possible tables and/or graphs that we could create?
  - We could create a graph for outcomes based on year, outcomes based on category, and category based on goals.

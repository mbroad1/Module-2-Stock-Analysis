# Module 2 VBA Challenge: Analyzing Stocks
---
## Overview of Project
### Background:
Steve is helping his parents in choosing the right green energy stocks to invest in by using data evaluating the stocks' performances in previous years. He believes that there are two factors to look for when deciding whether to invest in a stock: 1) a high yearly volume meaning that the stock is traded on a daily basis (the prices of stocks that are traded more frequently are thought to be more accurate than those of stocks traded rarely) and 2) the stock's price increased by the end of the year in comparison to the start of the year. Using VBA, we can help Steve evaluate these factors in two different fiscal years, 2017 and 2018, for twelve green energy stocks.
### Purpose:
The purpose of this analysis is to evaluate the yearly trends in stocks with code that is refactored so that it is more efficient and faster in executing the analysis. With the code that was originally made in Module 2, it is good for analyzing a relatively small number of stocks which means that if more stocks were to be analyzed, then the analysis would take much longer. Therefore, refactoring the code will reduce the analysis time and make the analysis more efficient to be used in the future. We will analyze the differences in timing to execute the code using the timer function in VBA on both modules created.

Likewise, we want to evaluate the differences in stock performances between 2017 and 2018 and see which stocks are steady in their increase and the stocks that are declining in price from 2017 to 2018. This information will help inform Steve and his parents on which green energy stocks are the best to invest in.

---
## Results
### Refactoring:
The code was refactored so that data from both 2017 and 2018 could be accessed without having to repeat the code made initially in the module.
#### yearValue Variable Example 1:
![Refactoring #1.1](https://github.com/mbroad1/stock-analysis/blob/main/VBA%20Challenge%20Resources/Refactoring%20%231.1.png)
#### yearValue Variable Example 2:
![Refactoring #1.2](https://github.com/mbroad1/stock-analysis/blob/main/VBA%20Challenge%20Resources/Refactoring%20%231.2.png)
As you can see in the above images of code, initializing the variable "yearValue" makes it so that the user can input either 2017 and 2018 to generate the analysis specifically for that year. Unlike this refactored code, the original code activated the 2018 worksheet so that all analyses were done with just that worksheet; therefore, in order to analyze the 2017 worksheet, we would have to copy and paste the code activating the 2017 worksheet for this piece of code. However, copying and pasting this code causes a longer analysis time.
#### Original Code Time for 2017:
![VBA_Challenge_Original_2017](https://github.com/mbroad1/stock-analysis/blob/main/VBA%20Challenge%20Resources/VBA_Challenge_2017_Original.png)
#### Original Code Time for 2018:
![VBA_Challenge_Original_2018](https://github.com/mbroad1/stock-analysis/blob/main/VBA%20Challenge%20Resources/VBA_Challenge_2018_Original.png)

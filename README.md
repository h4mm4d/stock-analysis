# Stock-Analysis
## Overview of Project
The objective of this project is to help Steve identifying best green energy stocks using VBA on MS Excel to help his parents make better informed decision on which stock is to be invested. This could be done by writing a script that:
1.	Summarizes the data in easily readable format. 
2.	Gives clear and meaningful picture of stock performance in a given year 
3.	Getting volume information on the stocks so get an idea of future trend. 
4.	Look at historical values to see how well certain stocks performed previously. 
The goal of this exercise is to produce reusable VBA script to see performance of stocks, if more worksheets for different years are added to this document. The data provided are two years’ worth of stock trading data on twelve different companies. That looks something like the example below. We are interested in finding only two points for each “tickers”, Total Daily Volumes and “Yearly Return”. “Total Daily Volume” can be calculated by adding all values for the ticker on the given years and “Yearly Return” is calculated by “(Final Close/ First Close)/1” for the year 
                   ##### Total Daily Volume (ticker) = Sum(Volumes)
                   ##### Return (%) = (Final Close/First close)-1 
  
## Results
It is important to understand what percentage historical return and volume means for future of trading. 
###Volume and Return Relations
	Volume means number of shares traded in a stock over a period. Volume is one of the indicators of market strength. Higher volume on rising price means market is excited about the stock and higher volume on falling price means market is still hopeful about the stock. On the contrary, falling price and lower volume means market is pessimistic about the given stock and rising price and lower volume means market is cautious about the given stock. 
### Analysis of the Stocks
As it shows from data that in year 2017 most green stock did well with exception of TERP and in 2018 none of the green stock market did very well except ENPH and RUN. The volume traded on those two stocks are exceptionally high as well. Therefore, it is advisable to invest money on those two stocks. FLSR and SEDG did well in 2017 however not so well in 2018. However, it seems market is still excited about those stocks. Therefore, FSLR and SEDG should be a buy as well. TERP, VSLR, and AY are sell. Rest of the stocks, including DQ, are mostly neutral base on these calculations.  

## Summary
VBA is a powerful language and easily available with MS Excel. It helped us calculate the stock performance over two years period. In this example where there were 3013 lines of data were neatly summarized in 12 lines of table with the help of VBA that made so much easier to classify the tickers based on performance. 

### Advantages of Refactoring
Advantage of refactoring is that the code works faster especially for a long set of data and it is better presented to the memory therefore faster calculation. Refactoring almost always improve the code and it is also good for people learning the code that can see where they are lacking in coding and then they can improve their coding practices. It also, keeps them from typing same code over and over again, hence not reinventing the wheels. Sometimes, programmer find easier solution to a complex that did not have to be complicated.
### Disadvantages of Refactoring
Sometimes refactoring codes take longer than rewriting, especially if they are not sufficiently commented. 


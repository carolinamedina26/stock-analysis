# Overview
## 1.0 Introduction
The Stock Market is a form of investment, in which financial brokers base their decision according to data analysis of stock performance to determine and predict their future value. In the following analysis, we’ll take a close look to different variables such as volume traded and return value of 12 stocks on the green energy industry and compare their performance during several years. 
## 1.1 Background
The goal of this project is to create a Macro that is able to provide the desired analysis of large datasets  in a short period of time, as well as,  a friendly interface for the user. For this reason, the initial code has been refactor to improve the time performance of the analysis 

# Analysis
## 2.0 Annual Analysis
The 2017 analysis displayed an overall positive return on most of the tickers, the ** higher return was 184.5%** on the stock ** SEDG ** and the lower positive return was  5.5% on the RUN ticker. The single ticker that display a **negative return is TERP with -7.2%**. (See graph) 
The 2018 analysis displayed an **overall negative** return on most of the tickers.  The ** only tickers  with positive returns were ENPH and RUN with 81.9 % and 84 % respectively.** 

## 2.1 Refactored Code
On the table above we can evidence how the refactoring reduced the waiting time of 
On the initial analysis ran in 0.4375 seconds for 2017 and , which was improve by refactoring the code. The way we improve the code was by the creating of a new variable `tickerIndex` to access the correct index. After this changes the code was simplified, for example: 

To obtain the `totalVolume` our code need it to have a conditional `ifhen` to located the tickers
`If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If`

After the new variable ‘tickerIndex` was introduce the task could be simplified by since the is part of the index. 

`tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(x, 8).Value`

# Summary 
## 3.0 Advantages and Disadvantages of Refactoring
Most of the time, the code is develop under circumstances where the data results are needed regardless of the efficiency of the code. For that reason, refactoring will be needed in a further step or second attempt to simplify the language, improve readability or  decrease the performance time of the analysis. On the other hand, more resources will need to be allocated to the solution of the same problem.  

## 3.1 Analysis of code
The advantages of the first draft of code is that was solved using several patterns, if is quick to implement but inefficient .On regards to the refactoring code. On regards to the refactor code is that is shorter and simplifies it, but the downs side is the extra time speeded to improve it. 






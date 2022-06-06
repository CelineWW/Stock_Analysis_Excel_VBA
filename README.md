# stock-analysis

##**OverView of Project**

###**Purpose**
1. To compare stock performance between 2017 and 2018.

2. To accelarate the execution time of all stock analysis for each year, the macro code was refactored. 

##**Result**

1. RUN is the stock worthy to invest.

![All Stocks Analysis 2017](https://user-images.githubusercontent.com/105877888/172074993-c71258d3-1b3a-4a69-a463-3c163e1d35ad.PNG)

![All Stocks Analysis 2018](https://user-images.githubusercontent.com/105877888/172074995-3c320e86-573b-45fa-bf8a-aee79ae3d46d.PNG)

  *Steve's parent planned to put their investment on DQ. Unfortunatly, DQ's return has plummeted. Absolutely, DQ is not a smart choice. From the view of return for 2018, the stock of ENPH and RUN got 80~85% return, which seems investable. However, Comparing to 2017, Enph's return dropped from 129.5% to 81.9%, RUN's return increased from 5.5% to 84.0%. Overall, ENPH is more like a promising profitbale stock to make investment.*

2. The execution of `All Stock Analysis` for each year(2017 & 2018) was accelarated. 

-- *The execution of `All Stock Analysis` for 2017 was sped up from `1.070313` to `0.1875` seconds.*

![Screenshot of run-time analysis for 2017](https://user-images.githubusercontent.com/105877888/172064893-62324114-946f-410e-913b-dfa34b8bfaaf.PNG)

![Screenshot of run-time analysis for 2017(Refactored)](https://user-images.githubusercontent.com/105877888/172064881-181bd289-a50a-4e76-a007-65d16378b380.PNG)

--*The execution time of `All Stock Analysis` for 2017 was sped up from `0.9609375` to `0.1875` seconds.*

![Screenshot of run-time analysis for 2018](https://user-images.githubusercontent.com/105877888/172064900-d4a6c153-6b26-40b5-88a7-4ce267c77ad6.PNG)

![Screenshot of run-time analysis for 2018(Refactored)](https://user-images.githubusercontent.com/105877888/172064885-18227c31-c1a5-4df9-9bc6-7455476c1b26.PNG)

##**Summary**

1. Two big changes were made under refactoring code.

--*Variable types of `startingPrice` and `endingPrice` were declared as `Single` instead of `Double`.* 

--*`TickerIndex` was introduced. This may avoid nested `For Loop`. *

2. Refactoring code was applied to the VBA script, it helps VBA script run 4~5 times faster.

--*Since Refactoring Macro could process starting Prices and endingPrice without decimal, the totalVolume can be caculated much faster.*  

--*Refactoring script contains 3 indepent for loops. Each row would be assigned to a certain tickerIndex. Thus each tickerIndex could loops over independently. The full worksheet could be only scanned for once.
Whereas in the orginal script with nested for loop. For Each tickers, all the rows in the worksheet would be scanned. So it would be loop over for 12 times in total.
This might be the main reason that execution time were damatically shortened. On the other hand, tickerIndex makes the code more complicate, which is easy to make mistakes for developer.*

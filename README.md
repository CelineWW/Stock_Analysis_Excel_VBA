# stock-analysis

##**OverView of Project**

###**Purpose**
1. To compare stock performance between 2017 and 2018.

2. To accelarate the execution time of all stock analysis for each year, the macro code was refactored. 

##**Result**

1. A user friendly year input box and clear button were created.
```
Dim yearValue As String
  yearValue = InputBox("What year would you like to run the analysis on?")
```
```
ClearWorksheet()
   Cells.Clear
   ```
2. RUN is the stock worthy to invest.

![All Stocks Analysis 2017](https://user-images.githubusercontent.com/105877888/172074993-c71258d3-1b3a-4a69-a463-3c163e1d35ad.PNG)

![All Stocks Analysis 2018](https://user-images.githubusercontent.com/105877888/172074995-3c320e86-573b-45fa-bf8a-aee79ae3d46d.PNG)

  *Steve's parent planned to put their investment on DQ. Unfortunatly, DQ's return has plummeted. Absolutely, DQ is not a smart choice. From the view of return for 2018, the stock of ENPH and RUN got 80~85% return, which seems investable. However, Comparing to 2017, Enph's return dropped from 129.5% to 81.9%, RUN's return increased from 5.5% to 84.0%. Overall, ENPH is more like a promising profitbale stock to make investment.*

3. The execution of `All Stock Analysis` for each year(2017 & 2018) was accelarated. 

-- *The execution of `All Stock Analysis` for 2017 was sped up from `1.070313` to `0.1875` seconds.*

![VBA_Challenge_2017](https://user-images.githubusercontent.com/105877888/172102290-c3fb1cc1-677f-4640-b836-5f74928c9b1a.PNG)

![VBA_Challenge_2017(Refactored)](https://user-images.githubusercontent.com/105877888/172102314-62d286a8-dae7-4970-8587-a64e62530e85.PNG)

--*The execution time of `All Stock Analysis` for 2017 was sped up from `0.9609375` to `0.1875` seconds.*

![VBA_Challenge_2018](https://user-images.githubusercontent.com/105877888/172102352-30c924b8-32be-42b9-b916-1608b6f67b25.PNG)

![VBA_Challenge_2018(Refactored)](https://user-images.githubusercontent.com/105877888/172102366-65073278-bcc1-496b-9e50-f0e7c69dd799.PNG)

##**Summary**

1. Two big changes were made under refactoring code.

--*Variable types of `startingPrice` and `endingPrice` were declared as `Single` instead of `Double`.* 

--*`TickerIndex` was introduced. This may avoid nested `For Loop`.*

2. Refactoring code was applied to the VBA script, it helps VBA script run 4~5 times faster.

--*Since Refactoring Macro could process `starting Prices` and `endingPrice` without decimal, `totalVolume` can be caculated much faster.*  

--*Refactoring script contains 3 indepent `For Loop`. Each row would be assigned to a certain `tickerIndex`. Thus each `tickerIndex` could loops over independently. The full worksheet could be only scanned for once.
Whereas in the orginal script with nested for loop. For Each tickers, all the rows in the worksheet would be scanned. So it would be loop over for 12 times in total.
This might be the main reason that execution time were damatically shortened. On the other hand, `tickerIndex` makes the code more complicate, which is easy to make mistakes for developer.*

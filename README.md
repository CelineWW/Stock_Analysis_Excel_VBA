# Stock Analysis Excel VBA


## Overview
1. To compare the performance of Wall Street Stock between 2017 and 2018.
2. Take user input for the year to show total daily volumn and return for each stock for specific year.
3. Create clear worksheet button to get ready for next analysis.
4. To accelerate the execution time of all stock analysis for each year, the macro code was refactored. 

## Result
1.  Run Stock Analysis
  - Sample code:
  
  ![Macro code example](https://user-images.githubusercontent.com/105877888/208582647-8e1af109-3c9c-4830-8191-2729b0a0b6c2.png)

  - Total daily volumn and return in 2017/2018 for each stock are displayed respectively on worksheet.

  <img width="1247" alt="2017 stock analysis" src="https://user-images.githubusercontent.com/105877888/208582462-3f1adfc2-9301-470b-9eac-7a684e4b353c.png">

  <img width="1247" alt="2018 stock analysis" src="https://user-images.githubusercontent.com/105877888/208582480-e196a1e0-4df9-407d-a581-a34e4c558f98.png">

2. User friendly year input box and Clear button
   - codes 
      ```
      Dim yearValue As String
        yearValue = InputBox("What year would you like to run the analysis on?")
      ```
      ```
      ClearWorksheet()
         Cells.Clear
      ```
   - Worksheet display
    <img width="1399" alt="take_user_input" src="https://user-images.githubusercontent.com/105877888/208582766-db2a0949-cc2b-4189-b164-0eabcd822a59.png">

2. Compare 2017 and 2018 analysis result to select the stock worthy to invest.

    ![All Stocks Analysis 2017](https://user-images.githubusercontent.com/105877888/172074993-c71258d3-1b3a-4a69-a463-3c163e1d35ad.PNG)

    ![All Stocks Analysis 2018](https://user-images.githubusercontent.com/105877888/172074995-3c320e86-573b-45fa-bf8a-aee79ae3d46d.PNG)


3. The execution of `All Stock Analysis` for each year(2017 & 2018) was accelarated. 

   - *The execution of `All Stock Analysis` for 2017 was sped up from `1.070313` to `0.1875` seconds.*

      ![VBA_Challenge_2017](https://user-images.githubusercontent.com/105877888/172102290-c3fb1cc1-677f-4640-b836-5f74928c9b1a.PNG)

      ![VBA_Challenge_2017(Refactored)](https://user-images.githubusercontent.com/105877888/172102314-62d286a8-dae7-4970-8587-a64e62530e85.PNG)

    - *The execution time of `All Stock Analysis` for 2017 was sped up from `0.9609375` to `0.1875` seconds.*

      ![VBA_Challenge_2018](https://user-images.githubusercontent.com/105877888/172102352-30c924b8-32be-42b9-b916-1608b6f67b25.PNG)

      ![VBA_Challenge_2018(Refactored)](https://user-images.githubusercontent.com/105877888/172102366-65073278-bcc1-496b-9e50-f0e7c69dd799.PNG)



## Summary

1. *Steve's parent planned to put their investment on DQ. Unfortunatly, DQ's return has plummeted. Absolutely, DQ is not a smart choice. From the view of return for 2018, the stock of ENPH and RUN got 80~85% return, which seems investable. However, Comparing to 2017, Enph's return dropped from 129.5% to 81.9%, RUN's return increased from 5.5% to 84.0%. Overall, ENPH is more like a promising profitbale stock to make investment.*

2. *Two big changes were made under refactoring code.*

- *Variable types of `startingPrice` and `endingPrice` were declared as `Single` instead of `Double`.* 

- *`TickerIndex` was introduced. This may avoid nested `For Loop`.*

3. *Refactoring code was applied to the VBA script, it helps VBA script run 4~5 times faster.*

- *Since Refactoring Macro could process `starting Prices` and `endingPrice` without decimal, `totalVolume` can be caculated much faster.*  

- *Refactoring script contains 3 indepent `For Loop`. Each row would be assigned to a certain `tickerIndex`. Thus each `tickerIndex` could loops over independently. The full worksheet could be only scanned for once.*
*Whereas in the orginal script with nested for loop. For Each tickers, all the rows in the worksheet would be scanned. So it would be loop over for 12 times in total.*
*This might be the main reason that execution time were damatically shortened. On the other hand, `tickerIndex` makes the code more complicate, which is easy to make mistakes for developer.*

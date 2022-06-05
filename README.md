# stock-analysis

##**OverView of Project**

###**Purpose**

To accelarate the runtime of all stock analysis for each year, the macro code was refactored. 

##**Result**

*1. Variable types of startingPrice and endingPrice were declared as Single instead of Double. So the macro could process starting Prices and endingPrice without decimal. *
```
Dim startingPrice As Double 
Dim endingPrice As Double
```   

*2. tickerIndex was introduced to avoid nested `For Loop`. So as that each tickerIndex could loops over independently. * 
```
For i = 0 To 11 
  ticker = tickers(i)
  totalVolume = 0
    Worksheets(yearValue).Activate
    
    For j = 2 To RowCount
      If Cells(j, 1) = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
      End If
      
      If Cells(j - 1, 1) <> ticker And Cells(j, 1) = ticker Then
        startingPrice = Cells(j, 6).Value
      End If
      
      If Cells(j, 1) = ticker And Cells(j + 1, 1) <> ticker Then
        endingPrice = Cells(j, 6).Value
      End If
Next j
```
--*Above is the Orginal VBA Macro with nested for loop. For Each tickers, all the rows in the worksheet would be scanned. So it would be loop over for 12 times in total. *

```
For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    For i = 2 To RowCount
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
  
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value    
            tickerIndex = tickerIndex + 1
        End If    
    Next i
```
--*Refactored VBA Macro contains 3 indepent for loops. Each row would be assigned to a certain tickerIndex. The worksheet could be only scanned for once. *

*2.The runtime of all stock Analysis for each year was accelarated. 
-- The runtime of all stock Analysis for 2017 was sped up from `1.070313` to `0.1875` seconds.*

![Screenshot of run-time analysis for 2017](https://user-images.githubusercontent.com/105877888/172064893-62324114-946f-410e-913b-dfa34b8bfaaf.PNG)

![Screenshot of run-time analysis for 2017(Refactored)](https://user-images.githubusercontent.com/105877888/172064881-181bd289-a50a-4e76-a007-65d16378b380.PNG)

--*the runtime of all stock Analyses for 2017 was sped up from `0.9609375` to `0.1875` seconds.*

![Screenshot of run-time analysis for 2018](https://user-images.githubusercontent.com/105877888/172064900-d4a6c153-6b26-40b5-88a7-4ce267c77ad6.PNG)

![Screenshot of run-time analysis for 2018(Refactored)](https://user-images.githubusercontent.com/105877888/172064885-18227c31-c1a5-4df9-9bc6-7455476c1b26.PNG)

##**Summary**
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
1. Original 

There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
2.


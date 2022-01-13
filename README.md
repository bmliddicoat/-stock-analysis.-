# Refactor VBA code

## Overview of Project
Refactor of Module 2 solution Code to determine if refactoring caused VBA script to perform more quickly.

### Explain the purpose of this analysis.

The client, Steve, was impressed with prior work done with the original workbook/stock analysis.  The client wants to expand the dataset to include entire stock market over the last few years.  The prior code worked well for twelve stocks but due to increased stock volumes, the original VBA Code may not be efficient and/or effective.  I must edit, refactor, the original code.  This will allow the code to loop through all the data one time in order to collect the same information.
  
## Results

### Stock Results of 2018 VS 2017
![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Screenshots/2017_Results.png?raw=true)

![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Screenshots/2018_Results.png?raw=true)

The stock results between 2018 and 2017 had quite varying returns.  The tables above help detail the stark differences between each year.  2017 had favorable return with all except TERP achieving positive gains.  Well, 2018, showed negative returns except for positive returns of RUN and ENPH.  In analyzing both years’ returns, ENPH shows the highest performance with a 129.5% return (2017) and 81.9% return (2018).  The clients original ask of looking into DQ seems favorable in 2017 return but large losses occurred in 2018.        


### Refractoring Code Performance

![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Resources/Org_Code_2017.png?raw=true)

![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Resources/Execution2017Refrac.png?raw=true)

![alt text](hhttps://github.com/bmliddicoat/Stock-Analysis/blob/main/Resources/Org_Code_2018.png?raw=true)

![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Resources/Execution2018Ref.png?raw=true)

In refactoring the original VBA code, execution times improved dramatically.  The images abovce help show the time cut down from refactoring.   The original code for 2017 ran at around a speed of .2578 seconds.  Well, the refactored code for 2017 was only .06445 seconds.  There was similiar result for 2018.  The orginal code had execution time of .2578 seconds.  The refactored of 2018 had a time of .0664 seconds.  This improvement is due to certain actions that occurred during the refactor coding.  

### VBA Script of Refactoring

![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Screenshots/All_Stock_Analysis1.png?raw=true)

![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Screenshots/Refactored1.png?raw=true)


Instead of using a nested loop through all rows until all data was analyzed as the original code, the refactored code created single loops to examine certain data sets and produce a result.  To achieve this, I took certain steps.  First, I had to set tickerIndex as Double that was equal to zero.  This was done to before the VBA code began to iterate all the rows.  The tickerIndex was created for index of the arrays that will follow this code.  The VBA code that was used for this first step was: Dim tickerIndex as Double tickerIndex = 0 .  The next step was to create single loop for the three output arrays.  These arrays were tickerVolumes, tickerStartingPrices and tickerEndingprices.  I had to include the index in paratheses after each array name.   I decided on eleven as value of the index due to starting at zero rather than one.  The final 12th will be represented by eleven.  The final VBA code is Dim tickerVolumes(11) As Long, Dim ticketStartingPrices(11) As Single, and Dim tickerEndingPrices(11) as Single. The third step was to begin a loop and to initialize the tickerVolumes to zero.  This was accomplished by starting the For loop as i = 0 to 11.  I then wrote tickerVolumes(i) = 0.  The i access the element of the array of tickerVolumes, accordingly.  After the code was written, it was vital to use the Next i line to ensure following code would use the following i rather than the prior.  The fourth step was to loop over all rows in the spreadsheet.

![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Screenshots/All_Stock_Analysis2.png?raw=true)

![alt text](https://github.com/bmliddicoat/Stock-Analysis/blob/main/Screenshots/Refactored2.png?raw=true)

I used the prior code from the original For i = 2 to RowCount.  The fifth step was to increase the volume of the current ticker rather than loop through rows to get total volume of a current ticker.  This was done by using the code  
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value .  Next, I checked the current row was the first row with selected tickerIndex with a if then statements. The code is If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value. Then, I checked if the current row is last row was selected.  Code to achieve this was If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then  tickerEndingPrices(tickerIndex) = Cells(i, 6).Value . These two if then statements ensure the cell in the current row with tickerIndex accoriding to either statment of being in first row or last row.   lastly, I increased the tickerIndex by if then statement. If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then  tickerIndex = tickerIndex + 1.  Last, was to output the date.  This was achieved by : For i = 0 To 11 
    Worksheets("All Stocks Analysis").Activate
    Cells(i + 4, 1).Value = tickers(i)
    Cells(i + 4, 2).Value = tickerVolumes(i)
    Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1 
    
In summary, the refactored code focused on row/cell relations and original code examined specific tickers.  This adjustment improved the execution times because the computer was able to examine all data within the worksheet rather than loop through the data for each stock.

## Summary

### Advantages or Disadvantages of Refactoring Code

Refactoring code can improve the efficiency of a code.  Another advantage is the code may look less complicated and “cleaner”.  This can help future users being able to work with the code.  The number one disadvantage of refactoring is time consumption.  In refactoring, a user may also not be able to know what a next step is to improve the code.

### Pros and Cons of Refactoring the Original VBA script

I was able to notice a vast improvement in execution time in refactored compared original VBA script.  In working with the script of the refactored script, I was also able to come back to it with and not get lost within a nested loop as the original.  Although, I ran into issues on determining what code I should use at certain points.  I had to reexamine the old code, search websites like stack overflow.  One example was having issues with loop to initialize the tickerVolumes to zero.  This process became time consuming for myself.  Overall, the refactored script is benefit to the client.  They will be able to implement the VBA Code for future data sets and have improved execution times.  This could be beneficial if the client has larger sets of stock data and years to examine.      

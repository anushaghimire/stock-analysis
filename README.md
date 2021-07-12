
# Overview of Project
## The purpose of the project
The purpose of this analysis is to edit or refactor the Module 2 solution code without changing the function but to make it efficient. Writing a fully working code a second time may give the programmer the luxury of time to make scripts efficient, fast, and with fewer memory hogs. In our exercise, we were able to perform the refactoring and made the VBA script run faster. 


# Results
## The analysis is well described with screenshots and code 
### Comparison between 2017 and 2018 stock performance
As we can see in the pictures, in 2017 most of the tickers are green while in 2018 most are red indicating the general market performance in those years. Long-term investors, mostly focusing on the Bull type of investments would have made a lot of money in 2017, while the Bear type of investment would have made money in 2018.  There are some exceptions, two tickers ENPH and RUN were solid green despite the market being red in 2018. Similarly, TERP was red in the year 2017 where all other tickers or market, in general, were green. 

<img width="337" alt="All Stocks_2017" src="https://user-images.githubusercontent.com/85364095/125206912-66793000-e23e-11eb-91fb-10e0092bb86d.png">  <img width="337" alt="All Stocks_2018" src="https://user-images.githubusercontent.com/85364095/125206914-68db8a00-e23e-11eb-84a4-96f940f6ed9a.png">

### comparison between Execution times of the original script vs refactored script
The run times for the original script for the years 2017 and 2018 are around 0.65 and 0.63 seconds and the run times for the refactored script are 0.11 and 0.125 seconds respectively. We can confidently say that refactoring the code did make the VBA script run faster, almost like 6 times faster than the old script. This improvement in run times does indicate new code is more efficient than the old one. 
#### Run times for the original script
<img width="400" alt="originalcode_VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85364095/125207000-e7382c00-e23e-11eb-8d23-9aba77fff79b.png"> <img width="400" alt="originalcode_VBA_Challenge_2018" src="https://user-images.githubusercontent.com/85364095/125207011-f1f2c100-e23e-11eb-9db4-2d36ddbb3a31.png">


#### Run time for the refactored script
<img width="400" alt="VBA_Challange_2018" src="https://user-images.githubusercontent.com/85364095/125207100-4a29c300-e23f-11eb-989a-4a534fef60ea.png">  <img width="400" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85364095/125207107-557cee80-e23f-11eb-98bc-9b807eefdd94.png">


### Refactored VBA code
Below is the refactored codes according to the Module 2 challenge instructions. Here I created a ticker Index and set it to zero then created three output arrays(tickerVolumes, tickerStartingPrices, and tickerEndingPrices). After that I created a "for" loop to initialize the ticker volume to zero and if the next row’s ticker doesn’t match, increase the tickerIndex. Then I created a "If then" statement to check if the current row is the first row with the selected tickerIndex, and if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the ticker StartingPrices and tickerEndingPrices variable respectively. Then wrote a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker. Finally created a "for" loop to loop through our 3 arrays(tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in my VBA_challenge spreadsheet. 





````
1a) Create a ticker Index
    
        tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row's ticker doesn't match, increase the tickerIndex
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
 '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
        End If
    
 Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerEndingPrices(i) - 1
        
Next i

````



## Summary
### The advantages and disadvantages of refactoring code in general.
#### Advantages:
- Refactoring the code improves readability of the code.  
- It makes it less complex and makes it easier to understand.
- It also helps to find bugs and makes programs run faster.
- The code will be a lot cleaner after refactoring.

#### Disadvantages
- It may take a long time to complete the refactoring process.
- Refactoring a given code while trying to make it short and efficient may introduce new bugs in the system.


### Advantages and disadvantages of the original and refactored VBA script.


The original VBA script contains nested loops. The refactored VBA script reduced the number of loops, which also reduced the run time for the script than the original script. 







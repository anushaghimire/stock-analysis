
# Overview of Project
## The purpose and background are well defined.
# Results
## The analysis is well described with screenshots and code 
### Comparison between 2017 and 2018 stock performance
<img width="337" alt="All Stocks_2017" src="https://user-images.githubusercontent.com/85364095/125206912-66793000-e23e-11eb-91fb-10e0092bb86d.png">  <img width="337" alt="All Stocks_2018" src="https://user-images.githubusercontent.com/85364095/125206914-68db8a00-e23e-11eb-84a4-96f940f6ed9a.png">

### comparison between Execution times of the original script vs refactored script
#### Run times for the original script
<img width="300" alt="originalcode_VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85364095/125207000-e7382c00-e23e-11eb-8d23-9aba77fff79b.png"> <img width="300" alt="originalcode_VBA_Challenge_2018" src="https://user-images.githubusercontent.com/85364095/125207011-f1f2c100-e23e-11eb-9db4-2d36ddbb3a31.png">
#### Run time for the refactored script
<img width="300" alt="VBA_Challange_2018" src="https://user-images.githubusercontent.com/85364095/125207100-4a29c300-e23f-11eb-989a-4a534fef60ea.png">  <img width="300" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85364095/125207107-557cee80-e23f-11eb-98bc-9b807eefdd94.png">








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
#### The advantages and disadvantages of refactoring code in general.
#### There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).







# stock-analysis
Module 2 Challenge

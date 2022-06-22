# VBA Stock Analysis
Module 2 Challenge

### Purpose
The challenge was to refactor VBA code that was created for a stock analysis for 12 stocks for both the year 2017 and 2018 during module 2 to make the code more efficient and decrease run time. 

## Results
### Analysis
To begin the refactoring process I took the original code that looped thourough each indivual ticker then displayed the output of each indivual ticker. By converting each indivual ticker to an array I was able to loop through all of the tickers then display the output of all the tickers at one time. As shown below. 

###Original Code

 'Loop through tickers
  For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0

       'loop through rows
       Worksheets(yearValue).Activate

       For j = 2 To RowCount
           'get total for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           'get startingPrice for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           'get endingPrice for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       'output for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

###Refactored Code
    
    1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i

       Worksheets(yearValue).Activate
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
         
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         End If
                
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
               
           
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
                
        Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
### Pros and Cons of Refactoring Code
Refactoring gives us a chance to clean the code so its more organized for a number of reason. Advantages of doing so are it allows for easier debugging it can improve the overall software allowing the code to run bettr and faster. It also allows for other programers to have a better understanding of the code. However there is not always the option to refactor code due to time constraints or the code could be thousands of lines or more and may pose a risk of breaking a function of the code thus increasing the time it would take to refactor all of the code. 

###Outcome of Refactoring
As you can see from the screen shots the original code for both 2017 and 2018 took over a second to run where the refactored code took about a quarter of the time to run. 

##Orginal 2017
![Stock Analysis 2017 Run Time](https://user-images.githubusercontent.com/106495422/175051929-322a9e48-1890-4da7-a60d-d901789a6940.png)

##Refactored 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/106495422/175052113-8e5058b6-39d4-45ad-a5d4-ceaf209eef30.png)

##Orginal 2018
![Stock Analysis 2018 Run Time](https://user-images.githubusercontent.com/106495422/175052181-13a8190a-9e9a-406f-9569-f3c8322fcf83.png)

##Refactored 2018
![VBA_Challenge_2018](https://user-images.githubusercontent.com/106495422/175052263-6ec8366b-4d1c-455f-8e27-d8de20f83e4b.png)

    

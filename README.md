# Module 2 | Assignment - Wall Street



# Exploring Energy Stocks for Green Returns




## Overview of Project
Steve,  has commission a data analysis of green energy stocks return for 2017 and 2018. We will use Excel and an VBA macro to perform the financial operation to analyze total ticker volume and total ticker percent return. 


## Analysis and Challenges
Using the provided  "Green_Stocks.xlm and challenge_starter_codes.vbs we executed a looping macro for the given dataset. The challenge asked us to include the entire stock market over the last few years for analysis. Additionally, the challenge asked us to refactor, the Module 2 solution code to loop through all the date one time. In essence using the provide starter codes, we were building this code for the first time. 

### Results
Using the macros for clearing the dataset and running the analysis provide a great tool that many clients will like. The ability to give that function directly in the excel document will provide for ease of use and increase functionality to all users. The "refactor" code used in the exercise return 0.09375 second for 2018 and 0.09375 seconds for 2017. The ability to run a large dataset and provide the users with quantifiable numbers and present a visual aid to tell the story of the data is extremely valuable. For the year 2017 Green Stocks, as a sector delivered positive returns. In the dataset 11 out of 12 analyzed stocks produced positive returns and 9 out 12 stocks produced double digits returns for 2017.  Using our stock analysis, we saw the year 2018 reverse gains and only 2 out of 12 stocks produced positive returns.  


The following is sample and excerpts of important code that was modified and recreated with the help of challenge_starter_codes.vbs, teaching assistances, and classmates. 

 '1a) Creating a variable, tickerIndex and setting it to 0
   For tickerIndex = 0 To 11

  '1b) Creating three output arrays for tickerVolumes, tickerStartingPrices, and tickerEndingPrices
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero. In addition to creating ticker Volume loop we are designing our pattern to utilized the mechanism of "For" loop and "Nest" loop to optimize or REFACTOR our code from the homework study. By nesting our loops the code is able to process one tickerIndex "stock ticker" via the same loop and analysis ticker name, total daily volume, and use tickerStartingPrices and tickerEndingPrices to produce precent returns. 


    Worksheets(yearValue).Activate

    For i = 0 To 11

      tickerVolumes(i) = 0

    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.

      For i = 2 To RowCount
      
      
        '3a) Increase volume for current ticker
             
             If Cells(i, 1).Value = tickers(tickerIndex) Then
              
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
            End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
                
                If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then

                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        

            '3d Increase the tickerIndex.
            
                tickerIndex = tickerIndex + 1
            
        End If

          
    Next i






### Summary, To Refactor or Not To Refactor
Refactoring code can have great benefits and advantages. Refactoring code can save you time in your initial development stages of writing code. The refactor code can run cleaner or smoother using less lines of code and computer resources. Often less line of code will produce better understanding and faster adaptation for the developer community.  

While every programmer would like to speed up operations using refactor code there are some disadvantages to refactoring codes.  Changing or refactoring the code to condense logical operation could have unintended result or output. Having fewer lines of code by nesting loop and analyzing the data fewer number of time is the goal of any developer. Realistically, it takes time and experience to implement the new code correctly. The ability to troubleshoot and debug a complicated code could be negated by the simplicity of a code with a few more lines. While every developer and programmer should try to adhere to the "Don't Repeat Yourself" rule, simple lines of code that have comments and a linear order flow will be easier to troubleshoot than code that have numerous orders of operations nested inside of loops, that are nest inside of other loops. 






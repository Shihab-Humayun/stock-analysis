# Kickstarting with Excel

## Overview of Project

### Purpose
The purpose of this project is to perform an analysis on the stock data sheet by using VBA macros. We would like to know the total daily volume and the return of each stock by making the user input a specific year. Although Steve admired the previous VBA script, which was able to display the stock analysis for a dozen of stocks, we would like to create a more efficient script. If there thousands of stocks needed to be analyzed, having an improved VBA script will be beneficial and take less time to calculate the data.

## Results

### Stock Performance
Majority of the stocks in 2017 had a positive return, which are highlighted in green. However, the stock TERP was the only one down by 7.2%. The highest stock return was from the stock DQ with a value of 199.4%. In 2018, majority of the stock returns went to the negatives, which are highlighted in red. The stock ENPH which is highlighted in green still had a decrease in return when comparing it to the year of 2017. The stock DQ took the lowest return percentage, which is recorded at 62.6%. Some of the total daily volume of the stocks from 2017 to 2018 decreased as well. DQ, ENPH, HASI, RUN, SEDG, TERP and VSLR are few stocks that have increased in total daily volume within that period of time. The following images displays the data of all the stocks in 2017 and 2019.

![image](https://user-images.githubusercontent.com/49353083/110181182-3ece1300-7dd9-11eb-8d38-9ee7ef4dd056.png)


![image](https://user-images.githubusercontent.com/49353083/110181043-edbe1f00-7dd8-11eb-8101-c27ff63f5a49.png)

### Execution Times
When analyzing the run time between the original script and the refactored script, we can see a decrease in time needed to compute the data on the refactored script. For the 2017 original script, it took around 0.816 seconds. The 2017 refactored script took around 0.168 seconds to calculate the same data, which is faster by around 0.648 seconds. The similar results are shown for the 2018 original script, which took around 0.828 seconds and 0.160 seconds for the 2018 refactored script. The following images shows the message dialog box of each script executed.

2017 Original Script

![image](https://user-images.githubusercontent.com/49353083/110181496-f2370780-7dd9-11eb-8076-b1b2c990c647.png)

2017 Refactored Script

![image](https://user-images.githubusercontent.com/49353083/110181920-4b06a000-7dda-11eb-9f88-c4d8e77f28e8.png)


2018 Original Script

![image](https://user-images.githubusercontent.com/49353083/110181597-27dbf080-7dda-11eb-81d5-8aff8de7a06f.png)


2018 Refactored Script

![image](https://user-images.githubusercontent.com/49353083/110181775-40e4a180-7dda-11eb-816d-8a92adb00133.png)


## Summary

### 1. What are the advantages or disadvantages of refactoring code?
The advantages of having the refactoring code is that it is able to execute the script and calculate the results needed from the worksheet quicker than the orignal script. The use of declaring the variables as arrays is able to run the script efficiently. The following variables were declared below:

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

Although it is able to compute the data from the worksheet faster, creating the code takes a little more time and thinking. For the most part, the original and refactored script has the same code layout. The main differences was using the tickers as an array and the other variables as well.

### 2. How do these pros and cons apply to refactoring the original VBA script?
The refactored script uses the same logic as the original script. The refactored script is beneficial because it will able to effectively analyze the stocks if Steve were to add thousands of stocks compared to the dozens. The original script will take a significantly longer time compared the the refactored version when calculating thousands of stocks on the worksheet.

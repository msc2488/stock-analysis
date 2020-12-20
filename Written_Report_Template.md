# Stock Analysis 

## Overview of Project
Steve is a recent college graduate with a finance degree and is helping his parents analyze green energy stocks. He is interested in tracking return and liquidity across his parents investable universe to determine where they should invest.

### Purpose
Help Steve automate his analysis using VBA to reduce the amount of manual work required and occurence of "human errors". Once the code was sufficienlty written, it was refactored to improve operating efficiency. To further improve the speed, the code was refactored a second time. Formatting was applied to make the results quickly interruptible.

## Results: General Market performance 
The analysis highlights that divergence in performance that often occurs between years. All of stocks analyzed had positive returns for the year in 2017, while in 2018 most had negative returns, as shown in the below images. This likely points to stronger overall equity markets in 2017 vs 2018 or cyclical pressures on the clean energy sector in 2018. For 2017, stock selection was not as important as it was in 2018. 

https://github.com/msc2488/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png
https://github.com/msc2488/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png

### Results: Refactored Stock Analysis 
Once the code was written and performing as expected, it was refactored to increase operating speed. To enable the code to run on any list of tickers, the tickerIndex was set to 0 using the code  tickerIndex = 0. Output arrays were created for volume, starting and ending prices, via 
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerendingPrices(12) As Single
This was refactored again using ReDim which further improved operating speed.
https://github.com/msc2488/stock-analysis/blob/main/Resources/VBA_Challenge_2017_ReDim.png
https://github.com/msc2488/stock-analysis/blob/main/Resources/VBA_Challenge_2018_ReDim.png

### Summary: General Refactoring
Refactoring code is a useful practice as it allows us to improve operating functionality and provides us with a starting point for our code. However because often the person working to refactor the code did not create it, there can be difficulty understanding the exisiting code if sufficient documentation isn't provided. Additionally not every change will improve functionality and some could in fact reduce functionality. 

### Summary: Refactored Code 
The refactored code reduced runtime and works for any data set vs only a previously defined array. However the original code is easier to understand for those new to coding and the code runtime was sufficient prior to being refactored, so the ultimate end user should be considered in this case. 


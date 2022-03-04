# VBA_Challenge
Performing analysis on the best stocks to invest in based on previous returns.
# Project Overview
- The Purpose of this project is to automate and analyze stock performances between 2017 and 2018 to make a clear choice in stock selection for best returns. 
# Results
- Based on observations regarding stock performance in 2017 vs 2018, the market in 2018 was doing exponentially worse than the 2017 year shown in the images below



![image](https://user-images.githubusercontent.com/99559096/156680504-0c5bd699-0348-4a3f-900c-296e2c5fc808.png)



![image](https://user-images.githubusercontent.com/99559096/156680582-d370b22c-257d-42ac-b927-00fc32e29193.png)


As shown, the 2018 year only contained two stocks that rose in value, compared to the 2017 year, both grew by a large margin. In the case of investing, the best choices for continued growth are the ENPH and RUN tickers. However, in the mind of an investor, one may prefer to purchase stocks based on their negative performance, with understandings of their previous performances and why there may have been a decay in returns.
## Timing comparison
- As far as original scripting goes, the primary difference is the lack of multiple arrays being formed to only loop through the data a single time. 


![image](https://user-images.githubusercontent.com/99559096/156684889-f40a010d-dca9-4555-84e8-0b6a897814cd.png)


The photo attached above is an example of the soul array in the original script. The issue with this is that the data sheet has to get looped through 11 times the extra data needed to populate "Total Daily Volume" and "Return"


![image](https://user-images.githubusercontent.com/99559096/156685042-aa1b1492-186c-4e5a-bce0-7229ca13ae97.png)


In the refactored script, the addition of three output arrays allowed for a more efficient data sweep, in that, the data had a place to be stored for the ability to be pulled later in the script for population on the analysis spreadsheet.

![image](https://user-images.githubusercontent.com/99559096/156685166-ef1b7674-812d-493f-8be7-6198096a5e1e.png)


Above, shows the output For loop, this portion populates the "All Stocks Analysis" spreadsheet with the data stored in the output arrays combined with the tickers array. This is a major advantage as far as the time efficiency goes. Shown Below are the comparisons of time between the Original scripting, and the Refactored scripting.


![image](https://user-images.githubusercontent.com/99559096/156685457-31ba51ce-8573-4317-b846-c24cfd22763d.png)
Original Scripting time

![image](https://user-images.githubusercontent.com/99559096/156685501-59ccc157-c13e-49cd-bc69-c67a443265b8.png)
Refactored Scripting time

This shows how the addition of output arrays allows the data to be looped through just 1 time. The efficiency of the Refactored Code as far as time goes, is quite literally 10x faster. 

# Summary
- Primary disadvantages of refactoring code is realistically the time it takes to re-code/reformat. Also, if the created program is already efficient for the job being performed, then refactoring for milliseconds of time is unneccesary. 
- Primary advantages : quicker reporting, clearer data organization inside the code as far as being able to read and understand what an array is assigned to. If there are hundreds of thousands of datapoints, then refactored code for the purpose may be neccessary for the program/computer to run in an effective amount of time. 

## Application to VBA
- In this project, refactoring the code allowed for a 10x more efficient reporting tool. The addition of arrays allowed for data to have a home to be pulled later in the scripting. 
- The Disadvantage to this refactoring, is that, although 10x more efficient, this is only mere milliseconds being saved. For the data being analyzed, there is no realistic situation that the milliseconds will cause issue. 
- Processing power of Excel may be limited by its engine, however, with refactoring, we show that it is clearly in the end user's hands to create an effective tool for the processing power of Excel.

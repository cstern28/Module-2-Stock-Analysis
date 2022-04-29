# Module-2-Stock-Analysis

## Overview
The purpose of the Stock Analysis project is to be able to create VBA code that performs the stock analysis for both years and then refactor the code to run faster than it did in Module 2.

## Results 
The original VBA code ran in about 1.05 seconds for years 2017 and 2018. The refactored code ran in about .36 seconds for 2017 and .35 for 2018. Refer to the following images for time it took the computer to perform stock analysis. The refactored code was faster because it used an index and arrays in the for loop. Instead of looping through all rows in the spreadsheet for every different ticker, it only looped through once and categorized the information based on different tickers. Refer to the original code for loop and the refactored code for the loop.

### Original Time - 2017
![image](https://github.com/cstern28/Module-2-Stock-Analysis/blob/main/Resources/Original_2017.png)

### Original Time - 2018
![image](https://github.com/cstern28/Module-2-Stock-Analysis/blob/main/Resources/Original_2018.png)

### Refactored Time - 2017
![image](https://github.com/cstern28/Module-2-Stock-Analysis/blob/main/Resources/Refactored_2017.png)

### Refactored Time - 2018
![image](https://github.com/cstern28/Module-2-Stock-Analysis/blob/main/Resources/Refactored_2018.png)

### Original - For Loop
![image](https://github.com/cstern28/Module-2-Stock-Analysis/blob/main/Resources/Original_ForLoop.png)

### Refactored - For Loop
![image](https://github.com/cstern28/Module-2-Stock-Analysis/blob/main/Resources/Refactored_ForLoop.png)

## Summary 

### Refactoring Code in General
The advantages of refactoring code is that is helps the code run more efficient and uses less computer time to process the code. It also helps someone understand the code better if the code is written in a better, more efficient process and will be easier to follow. The disadvantages of refactoring code is that it may be time consuming to figure out how to make the code faster and written in a more concise, efficient manner. Another disadvantage is that it may also confuse the coder and not get the code to produce the same results if done incorrectly.

### Original and Refactored VBA Script
The advantage of refactored VBA script for stock analysis is that it runs faster and more efficient. The disadvantage was I had to figure out how to use indexing and arrays successully, which I eventually managed with help from a classmate. Also, it's important to note that the script runs based on tickers as defined, but the script would be difficult to run if the list of tickers was much larger. It would not be efficient to list and define thousands of tickers out by name.


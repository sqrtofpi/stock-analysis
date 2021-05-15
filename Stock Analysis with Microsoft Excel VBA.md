# Stock Analysis with Microsoft Excel VBA

Microsoft Excel Macro-Enabled Stock Analysis file link: [Module 2 Challenge - VBA](https://github.com/sqrtofpi/stock-analysis/blob/1b4d486168592de63e9ba1146d2e130f44895014/VBA_Challenge.xlsm)

## Project Overview

Having been commissioned by our client Steve to analyze stock performace for a company named DAQO to help his parents decide whether or not to invest, Steve has re-enlisted our help to broaden the original scope of his request. Steve "loves" the workbook we originally created for him and is now seeking updates to the spreadsheet to enable analysis of the entire dataset encompassing the entire stock market over the 2017 and 2018 timeframe. This project will refactor the original VBA code in Microsoft Excel to meet Steve's objectives and also improve the performance of the scripts.

## Methods
The refactoring and performance evaluations analysis were performed by Scott MacDonald on May 15th, 2021. Using already created "challenge_starter_code.vbs", the code was copied and pasted into the already existing file named "VBA_Challenge.xlsm" which is hyperlinked at the top of this analysis for convenience. The following modifications were made to the existing spreadsheet:

- Created module 4 with the macro name - **AllStocksAnalysisRefactored**

- Added line of code called tickerIndex to improve speed of the script by decreasing the number of loops performed:

  ![](../../Module%201%20-%20Kickstarting%20with%20Excel/Module%201%20Challenge/Resources/VBA_Challenge_tickerIndex.png)

- Created additional buttons to easily compare original VBA script vs. refactored VBA script

  ![](../../Module%201%20-%20Kickstarting%20with%20Excel/Module%201%20Challenge/Resources/VBA_Challenge_buttons.png)

## Results

When the VBA script was refactored, success of the project was gauged using 2 criteria:

1. Are the outputs the same from the original script to the refactored script?
2. How much faster/slower was the script from the original to the refactored code?

### 2017 Data

#### Original Results

![](../../Module%201%20-%20Kickstarting%20with%20Excel/Module%201%20Challenge/Resources/VBA_Challenge_2017_Original.png)

#### Refactored Results

![](../../Module%201%20-%20Kickstarting%20with%20Excel/Module%201%20Challenge/Resources/VBA_Challenge_2017.png)

### 2018 Data

#### Original Results

![](../../Module%201%20-%20Kickstarting%20with%20Excel/Module%201%20Challenge/Resources/VBA_Challenge_2018_Original.png)

#### Refactored Results

![](../../Module%201%20-%20Kickstarting%20with%20Excel/Module%201%20Challenge/Resources/VBA_Challenge_2018.png)

### Analysis of Results

For both measures of success, refactoring the code maintained the accuracy of the original code and improved the time it took to execute. Utilizing Microsoft Excel for Mac (version 16.49) with Microsoft Visual Basic for Applications (version 7.1) on a MacBook Pro running macOS Big Sur (version 11.2.3) with an M1 chip yielded the following results:

**2017 improvement = 73.1% reduction in code execution time** (from 0.3046875 seconds to 0.08203125 seconds)

**2018 improvement = 71.8% reduction in code execution time** (from 0.3046875 seconds to 0.0859375 seconds)

## Summary

- **What are the advantages or disadvantages of refactoring code?**

  **Advantages:** The obvious advantages of refactoring code are the performance enhancements. Over time, new programming languages and methods advance that make older ways of performing the same task obsolete or ineffective. This is most commonly seen with major software developers updating their operating system code to improve it's performance and add new features. Code should be evaluated periodically based upon how often it is used and the risk vs. reward of refactoring the code. Sometimes it is necessary due to misunderstood effects of standardized coding such as in the case of the "Y2K" coding issue which caused concern and many programs to be refactored / updated.

  **Disadvantages:** The primary disadvantage of refactoring code is that it could have unintended consequences for other programs or hardware that could potentially cause it to fault out or crash. There is also the cost associated with going through many lines of code and assuring the dependencies are not causing other unanticipated problems.

- **How do these pros and cons apply to refactoring the original VBA script?**
  With refactoring the original VBA script in this project, there was little risk as the code was only 131 lines of code in the refactored script. It is an isolated subroutine that only has a specific function meaning there are no other dependencies to consider and the time involved was not substantial and should yield time savings everytime it is used and especially if the dataset it much larger. 
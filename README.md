# Using Excel VBA to Analyze Stock Performances

## Overview of Project

### Purpose

The purpose of this project was to refactor Excel VBA code that performs a stock analysis specific stocks during the years 2017 & 2018. The goal for this project was to increase efficiency by improving the overall logic of the code and make it easier for future users to read as wells as understand which stocks performed better from 2017 to 2018.

### The Data

The data used to perform the analysis consists of 12 specific stocks during the years 2017 & 2018. Each year's data is separately displayed on their own respective worksheets and the data contains stock tickers, their daily trading volume, and various daily stock price metrics. **I wrote a code that identifies each _stock ticker_ and calculates the _annual return percentage_ and _total daily trading volume_**.

### The Results

#### Code Improvemnt 

I measured the refactored code performance by adding a script that calculates how long the code takes to identify each stock and calculate the **annual return percentage** & **total daily trading volume**. Once fully executed, the script outputs the elapsed time in a message box. Below are images of the message box displaying the time results.

***Original Code Time Results - 2017***

![Original_ScriptTime_2017](https://user-images.githubusercontent.com/107579508/176333879-f971ec04-159b-4590-b8d2-05b582d7007c.png)

***Original Code Time Results - 2018***

![Original_ScriptTime_2018](https://user-images.githubusercontent.com/107579508/176334437-27b7925d-f5a3-400a-8e53-d32642fbcfa5.png)

***Refactored Code Time Results - 2017***

![VBA_Challenge 2017_png](https://user-images.githubusercontent.com/107579508/176335887-970d9166-af47-4ba4-a1ff-3ac6b9454dab.png)

***Refactored Code Time Results - 2018***

![VBA_Challenge_2018](https://user-images.githubusercontent.com/107579508/176336079-528c801b-61da-44bb-9f90-19f09f753fce.png)

#### Stock Anaysis

As you can see below, 11 out of 12 stocks provided a positive annual return in 2017 while only 2 out of the 12 stocks provided a positive annual return in 2018.

Stock Analysis Results - 2017

![Stock_Performance_2017](https://user-images.githubusercontent.com/107579508/176337439-b5a96f13-5d25-4593-af9c-f643ee76b41d.png)

Stock Analysis Results - 2018

![Stock_Performance_2018](https://user-images.githubusercontent.com/107579508/176337539-46e2d057-fe9a-4dbe-8a8d-453681da384c.png)

### Summary - Advantages & Disadvantages of Refactoring

In general, refactoring code is a key part of the coding process. Its' general advantages are more efficient code, using less memory, and improving code logic to make it easier for future users to understand. However, it can often be time consuming and involve multiple developers. In a work environment that is demanding and fast-paced, refactoring complex code could introduce new bugs/errors and ultimately undermine efficiency.  

#### Refactoring the Stock Analysis Code

Refactoring helped me solidify my understanding of important VBA concepts and pushed me to think about solving a problem in a different way. By the end of this project's code was faster and cleaner for any user to understand. In my opinion few disadvantages werediscovered. One arguement could be made in regards to how much time I spent refactoring and the overall improvment in speed of computation. Are a few hours of time worth saving approximately .6875 seconds?

Overall, I still believe the advantages of refactoring code outweighed any disadvantage I could devise.



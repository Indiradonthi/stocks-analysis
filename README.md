### OVERVIEW: VBA Stock Analysis Project

Purpose
The purpose of this project is to help Steve analyze data in return to help his parents to make right choices to invest on Green energy stocks. In this project we are helping Steve look into DAQO stock and diversify the funds. 
Analysis and Challenges
Here's a quick look at the Stock Analysis and Challenges of this Project, including the following tasks:
•	Prepare our data set VBA_Challenge.vbs file for the project.
•	Create a resources folder in GitHub to hold the run-time pop-up messages that we’ll screenshot after running refactored analyses for 2017 and 2018.
•	Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
•	Follow the steps listed to Refactor VBA code and measure performance.
Our Challenge Data Background
Using the data set we will analyze the Green energy stock by building automated tasks. We will edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order, to collect an entire dataset. Then, we’ll determine whether refactoring the code successfully made the VBA script run faster. Finally, we just want to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
Steve wants to find the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. 
First, we had to calculate the DQ's total daily volume, we needed to loop through all the stocks, so we've typed the number of rows into the code itself. What would be even better, though, is to use VBA to find the number of rows to loop over. Unfortunately, VBA doesn't have a nice function or method to figure that out. 
Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. 
RESULTS: Refactor VBA Code and Measure Performance
Steve loves the workbook we prepared for him. At the click of a button, he can analyze an entire dataset. We were able to clearly see by refactoring code it was more efficient in running. To perform his analysis on larger datasets we created the Timer function displaying the Elapsed Time by assigning the macro. To run different types of stocks in the future created flexible macro for running multiple stocks. We formatted the results to understand at a quick glance with numeric formatting, apply bold, color, percent, decimal point, and borders.
Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:
1.	The tickerIndex is set equal to zero before looping over the rows.

Created a ticker Index variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requirement.
  
2.	Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a Long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.
 


3.	The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.
 
4.	 The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
 




5.	Stored values from tickerStartingPrices and tickerEndingPrices
Created an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable.
 











6.	Code for formatting the cells in the spreadsheet is working.
We made positive returns are color coded green and negative returns red, to be a lot easier to determine which stocks did well and which ones did not. Added some formatting based on the values of the returns by highlighting, applied bold, color coded the heading and applied the number format.
 
7.	Adding Comments is a Best Practices for Writing Super Readable Code
•	Commenting & Documentation
•	Consistent Indentation
•	Avoid Obvious Comments
•	Code Grouping
•	Consistent Naming Scheme
•	Avoid Deep Nesting,
•	Keeping the comment short
 








8.	The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

Finally, we run the stock analysis, to confirm that our stock analysis outputs for 2017 and 2018 are the same as dataset example provided. 

Final VBA Analysis 2017
2017 before refactoring:
 

2017 After refactoring:

 


Final VBA Analysis 2018

2018 Before refactoring:
 

2018 After refactoring:
 






SUMMARY: Our Statement:
Deliverable with detail analysis:
1. What are the advantages or disadvantages of refactoring code?
We can perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.
Advantages:
•	Logical errors easily appear in well structure code that contains nested conditionals and loops.
•	In our case, using Excel flow displays program logic in a more applicable manner, not tied to the order that the underlying code is written.
•	VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.
Disadvantages:
•	Refactoring process can affect the testing outcomes.
•	A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
•	A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
•	A complex unstructured code is usually best to split in several functions.
2. How do these pros and cons apply to refactoring the original VBA script?
	Advantages
a.	Refactoring is a good weapon to maintain the code 
b.	It's interesting thing to do whether part of current task or as a separate task 
c.	 Make the code clean and organized 4. Help to follow principles like SOLID, GRASP, etc
Disadvantages:
a.	It's risky when the application is big.
b.	It's risky when the existing code doesn't have proper test cases.
c.	It's risky when developers do not understand what's all about

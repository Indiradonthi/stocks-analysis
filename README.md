# **OVERVIEW: VBA Stock Analysis Project**

## **Purpose**

The purpose of this project is to help Steve analyze data in return to help his parents to make right choices to invest on Green energy stocks. In this project we are helping Steve look into DAQO stock and diversify the funds. 

## **Analysis and Challenges**

Here's a quick look at the Stock Analysis and Challenges of this Project, including the following tasks:

•	Prepare our data set VBA_Challenge.vbs file for the project.

•	Create a resources folder in GitHub to hold the run-time pop-up messages that we’ll screenshot after running refactored analyses for 2017 and 2018.

•	Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.

•	Follow the steps listed to Refactor VBA code and measure performance.

## **Project Data Background**

Using the data set we will analyze the Green energy stock by building automated tasks. We will edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order, to collect an entire dataset. Then, we’ll determine whether refactoring the code successfully made the VBA script run faster. Finally, we just want to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

Steve wants to find the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. 

First, we had to calculate the DQ's total daily volume, we needed to loop through all the stocks, so we've typed the number of rows into the code itself. What would be even better, though, is to use VBA to find the number of rows to loop over. Unfortunately, VBA doesn't have a nice function or method to figure that out. 

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. 

## **RESULTS: Refactor VBA Code and Measure Performance**

Steve loves the workbook we prepared for him. At the click of a button, he can analyze an entire dataset. We were able to clearly see by refactoring code it was more efficient in running. To perform his analysis on larger datasets we created the Timer function displaying the Elapsed Time by assigning the macro. To run different types of stocks in the future created flexible macro for running multiple stocks. We formatted the results to understand at a quick glance with numeric formatting, apply bold, color, percent, decimal point, and borders.

## **Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:**

**1.	The tickerIndex is set equal to zero before looping over the rows.**

Created a ticker Index variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requirement.


  ![1](https://user-images.githubusercontent.com/90879122/136734181-121cfc55-f1c9-4e52-9809-f5b7f9b3ade5.png)

**2.	Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.**

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a Long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.
 

![2](https://user-images.githubusercontent.com/90879122/136734345-26b210e0-f542-46ee-a171-d6062a5a9347.png)

**3.	The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.**

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.
 
 
 ![3](https://user-images.githubusercontent.com/90879122/136734450-ed186dbf-18e1-415a-83f5-438c8030f84d.png)

**4.	 The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.**

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
 

![4](https://user-images.githubusercontent.com/90879122/136734505-5571dc32-e22f-49e5-bc82-1869a191cd51.png)



**5.	Stored values from tickerStartingPrices and tickerEndingPrices
Created an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable.**
 


![image](https://user-images.githubusercontent.com/90879122/136734550-1c83f629-baf9-418c-b83f-320bd91c2755.png)



**6.	Formatting gives a better vision**

We made positive returns are color coded green and negative returns red, to be a lot easier to determine which stocks did well and which ones did not. Added some formatting based on the values of the returns by highlighting, applied bold, color coded the heading and applied the number format.


 ![6](https://user-images.githubusercontent.com/90879122/136734624-f46eb668-5212-4f32-a81b-e8f2a8b03e3b.png)

**7.	Adding Comments is a Best Practices for Writing Super Readable Code**

•	Commenting & Documentation
•	Consistent Indentation
•	Avoid Obvious Comments
•	Code Grouping
•	Consistent Naming Scheme
•	Avoid Deep Nesting,
•	Keeping the comment short

![7](https://user-images.githubusercontent.com/90879122/136734706-64c9e61e-6610-4749-98a9-66769bbeb4c1.png)


 










## **Final VBA Analysis for 2017**

**2017 Before refactoring:**


![8](https://user-images.githubusercontent.com/90879122/136734770-a16573df-f11b-49f7-b98d-9cdc1c5dec03.png)

**2017 After refactoring:**

 
![9](https://user-images.githubusercontent.com/90879122/136734819-3bd4e667-14cb-409e-8040-a3cf69417b3d.png)


**Final VBA Analysis for 2018**

**2018 Before refactoring:**

 ![10](https://user-images.githubusercontent.com/90879122/136734898-5275ebd5-39f8-43a5-b489-7672bd9784eb.png)


**2018 After refactoring:**
 


![11](https://user-images.githubusercontent.com/90879122/136734932-24666e2e-9a89-412a-9cd2-a9aca874e4da.png)





## **SUMMARY**

## **Deliverable with detail analysis:**

**1. What are the advantages or disadvantages of refactoring code?**

We can perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

**Advantages:**

•	Logical errors easily appear in well structure code that contains nested conditionals and loops.
•	In our case, using Excel flow displays program logic in a more applicable manner, not tied to the order that the underlying code is written.
•	VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.

**Disadvantages:**

•	Refactoring process can affect the testing outcomes.
•	A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
•	A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called         from the other functions.
•	A complex unstructured code is usually best to split in several functions.

## **2. How do these pros and cons apply to refactoring the original VBA script?**

Advantages
	
a.	Refactoring is a good weapon to maintain the code 
b.	It's interesting thing to do whether part of current task or as a separate task 
c.	 Make the code clean and organized 4. Help to follow principles like SOLID, GRASP, etc

Disadvantages:
	
a.	It's risky when the application is big.
b.	It's risky when the existing code doesn't have proper test cases.
c.	It's risky when developers do not understand what's all about

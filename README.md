# VBA of Wallstreet

## Overview of Project
This challenge is part of module 2 in the Data Science Bootcamp.
The purpose of this project is to develop experience with VBA in Excel.
In this challenge, we were tasked with refactoring code we had previously developed as part of the learning modules.

## Results
The original code would read through all the rows in a spreadsheet looking for specific data at each iteration.
The refactored code only needed to go through each row once, and took advantage of arrays to store specific data while reading each row.
We wanted to compare the output of the original code with the refactored code to ensure they were the same, and to seeif there were any 
improvements in run-time.

Based on the two figures below, the run-time for the 2017 refactored code is approximately 0.073 seconds.  
By comparison, the original code took 0.45 seconds to run.
In both the original and refactored 2017 codes, the DQ and SEDG stocks had the greatest positive returns of 199% and 185% respectively.

The run-time for the 2018 refactored code is 0.079 seconds.
By comparison, the original code took 0.53 seconds to run.
In both the original and refactored 2018 codes, the ENPH and RUN stocks had the greatest positive returns of 84% and 81% respectively. 

![Graph](/Resources/VBA_Challenge_2017.PNG)
![Graph](/Resources/VBA_Challenge_2018.PNG)

### Summary

The advantages of refactoring code include the ability to compile and run much faster than the original versions.  We are able to make more efficient
code that will be easier for someone else to analyze and read.  But the disadvantage of the refactoring we did was that we needed more memory as
arrays were used as an integral part of the analysis.  

With respect to the original VBA script, we certainly were more efficient in the refactored code.  And because of the efficiency improvements (more loops), 
it would be easier for someone to read and debug the analysis.

However, the code at the onset still defined the array tickers with a fixed value of 12.  If anyone were to change the excel spreadsheet by adding or deleting
ticker symbols, the code would be adversly affected.  Similarly, if the 2017 and 2018 spreadsheets were not ordered alphabetically by Ticker name, and then by date, 
there could be issues with both the original and refactored code. 

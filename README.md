# stock-analysis
VBA Challenge


Overview of Project

The purpose of this assignment is to refactor code for Steve and his parents. While we worked on over a dozen stocks in the learning module, in this challenge, we’ll be looking at doing more research on the stock market and ensuring it is easy to comprehend. By the end of the analysis, we aim to run all the stocks in a short amount of time.

What is refactoring? According to the dictionary, refactor is to restructure (the source code of an application or piece of software) so as to improve operation without altering functionality. In this analysis, we’ll analyze the steps we took to refactor Steve’s workbook, challenges encountered and some of the advantages and disadvantages of refactoring. 

Results

To get the desired results, we created a tickerIndex Variable, set it equal to zero and reiterate it over all the rows. To access the correct index, we created three output arrays (tickerVolumes, tickerStartingPrices and tickerEndingPrices). Then, we created a for loop to loop over all the rows in the workbook. To initialize the array of tickers, we created the 12 tickers as String, activated the data worksheet, looped over the rows, looped through the arrays, formatted the worksheet before running the code to get the result. We ran the code for both 2017 and 2018 years.

For 2017, we can see that it took 0.0625 seconds to run all the stocks for 2017 while it took 0.0703125 seconds to run analysis on the stocks for 2018.

<img width="298" alt="image" src="https://user-images.githubusercontent.com/85265504/124414137-b2e0df00-dd17-11eb-8ea9-94c38da8d8f1.png">

<img width="283" alt="image" src="https://user-images.githubusercontent.com/85265504/124414152-bc6a4700-dd17-11eb-804a-22191bcadfb2.png">

 

CHALLENGES

VBA is quite challenging, and I can see why it is not a favorite topic for many. During this challenge and in the class module, I encountered several challenges as even the smallest error, such as incorrect indentation can generate an error and it might take a while to find the error because we are unable to run codes until we finish typing the codes. Accidentally typing proces instead of prices generated an error, typing cells (i, 3) instead of (j, 3) also generates an error. It is incredibly time consuming and at times the information entered can generate the wrong result. For instance, during this challenge, when I ran the macro for 2017 and 2018, I got the exact same result. Not the same time, but the information in the Return column was the same and I know that is false. I had issues fixing the error as I couldn’t figure out why the numbers are the same when they should differ and I went over this for hours. After hours of analyzing, I reached out to the Help Desk on Slack and while we were going over the button, I realized it was the button I created for the All Stock Analysis that caused the error. After deleting the button, the code ran as it should and I created another button for the refactored analysis. VBA is definitely not a favorite of mine but with practice, I believe it’ll get easier. Like the saying goes, “Practice makes Perfect”.


Summary

 According to https://aip.scitation.org/doi/abs/10.1063/1.3516393?journalCode=apc, Refactoring improves the design of software, makes software easier to understand, and helps us find debugs and also helps in executing the program faster. Based on module and assignment, I find that statement to be true. I had to debug several times during this assignment but by the end of the assignment, the time to run all the stocks definitely decreased. Once you understand how the codes and system works, it definitely makes things easier and faster. The downside about refactoring is that it is very time consuming, and I definitely did not expect it to take as much time as it took to complete the coding and debug the information. If another system was utilized, it will definitely take longer that 0.0625 or 0.0703125 seconds to run the amount of data in the 2017 and 2018 worksheets. While this was definitely a learning experience for me, it is one I cherish and hope to get better at. 

# stocks_analysis

## Project Overview

-In this project I wrote a code in VBA in which I edited an existing code of mine to run smoother and more efficiently. The purpose of this code is to analyze a list of stocks for a client named Steve, and to calculate which of these stocks had a positive or negative return. My original code worked correctly, however once I refactored the code it appears more clean and runs quicker. Upon running the code, the user is asked which year they want to analyze--either 2017 or 2018. Once the year is typed, a message box is displayed on the screen which tells the user how quickly the code was executed. This data changes depending on which year the user decides to run. When viewing the Excel worksheet after running my code, one is able to see the return of all 12 tickers and whether or not they are positive for that year (green) or negative (red).

## Project Results

- After writing this code I have come to the result that 2017 was a much better year for our 12 stocks being analyzed. 11 out of the 12 stocks (or 92%) have a positive return value for this year, which is a big difference when comparing this to the data in our 2018 worksheet. In 2018 only 2 out of 12 (or 17%) of our stocks had a positive return value.

Please view: twenty_seventeen.png
	     twenty_eighteen.png

For photos of these results.

-Also attached are two images of the different execution times for each year.
The code for 2017 had an execution time of .1016 seconds and the code for 2018 had an execution time of .1172 seconds. These execution times are faster than those in my original code because I had refactored the code to run smoother and faster the second time around.

-In my original code, the code for the year 2017 ran in .2656 seconds and in 2018 it ran in .2539 seconds.

-This is proof that refactoring your code can yield better and faster results. Although a code may successfully run, there are instances in which they can always be improved.

##Summary

-When it comes to refactoring code there are both advantages and disadvantages to take into consideration. 

-One advantage is that we can improve the code's performance as well as aesthetics when we refactor. By removing clutter in code, or by reducing our code to fewer lines (making sure that the code runs the way we want it to) we can modify our code to perform faster and more efficiently. Also if someone were to read our code, we want to make sure that it's easy on the eyes and easy to interpret. It's always possible that we can refactor our code to make it more clean and legible. 

-Some disadvantages to refactoring code could be that we have spent more time on the code, therefore we are putting in more work which can be tedious. In order to be good coders, however, we must be willing to work hard and put more time into projects. Also, we could be trying to refactor the wrong lines of code giving us an error. Perhaps we try to rewrite a line of code that can't or shouldn't be rewritten and is already in its simplest and most efficient form. 

-The pros I have mentioned apply to this specific VBA script that I have refactored in that we have evidence that the code was executed faster. I can run both the original code and the refactored code and see that the original was slower and therefore inferior to the second one. The use of the comments broken down in the refactored code also make it much easier to comprehend if someone else were to be analyzing our code.


-The disadvantages of refactoring my VBA script are that I had to think much harder and put in more time and effort to make sure the code ran faster. I came across several errors and had to tweak many of my lines in order to get the code to where it is now. The hard work and the brainstorming paid off, however, as I was able to refactor my VBA script to perform better. 

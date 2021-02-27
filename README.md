# Annual Stock Analysis for 2017/8


## Overview of Project

This analysis reviewed annual growth in selected stocks over a 2 year period.  Each stock had daily open and closing prices as well as volume traded.   This basic data was used to identify year over year valuations.
The project also offered a way to understand loops in code and the ability to refactor code for better efficiency.

### Purpose

The client would like to understand stock growth over the course of each year reviewed.

## Analysis and Challenges

The initial looping sequence was clunky and slow.    Learning to do it a bit leaner and more quickly was a challenge.  Still not sure how I got here.   I did not get good results in getting the code to run until I copied the code to a notepad and recopied it into a fresh spreadsheet.  It worked immediately following that.  Wasted 6 hours tinkering with code trying to get it to run.   I was not a happy camper.   

## Results

Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

The Results of the analysis were very clear.  Annual growth in 2017 was very good and in 2018, they were much poorer across the board.
- [VBA_Challenge_2017.png](VBA_Challenge_2017.png)
- [VBA_Challenge_2018.png](VBA_Challenge_2018.png)


## Summary: In a summary statement, address the following questions.

This code will review the data and pull the start and end pricing over a given data set as long as it is in date order.   
To review the VBA code and macro, or the data, please see the link below:

-[VBA_Challenge.xlsm](VBA_Challenge.xlsm)

### What are the advantages or disadvantages of refactoring code?

Refactoring code has many advantages.  There is almost always a better way of doing something.  Given the chance to reflect the ability to revise and retune code or “refactor” so it is more efficient is an excellent way to continue to be able to build more and more sophisticated code and reduce the chance of error as one adds features and complexity.

There are few disadvantages to refactoring code.  One critical issue is that often the person who is refactoring is not the person who wrote the code, so the person refactoring must have a very clear understanding of what the functionality AND intent of the code actually is.  Long term stability may be degraded by refactoring code long after the product has been introduced.   Long term users may be upset by slightly different ways of doing things and so the user base declines.

### How do these pros and cons apply to refactoring the original VBA script?

In refactoring the code, I learned a better way to do loops.  Still don’t quite understand it but its supposedly a better way.   The code did run faster.  And the microseconds it shaved would be important if I was running hundreds of different tickers.  However, this code is already faulty as it needs the data to be in a specific order so I doubt making it better in this context will much affect its usability.

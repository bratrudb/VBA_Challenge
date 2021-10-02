# VBA Challenge
Module 2 Assesment
## Overview of Project
In Module 2 we were given stock performance data for twelve companies for the years 2017 and 2018. We initally started our analysis looking at one company in particular (DAQO) and found their return in 2018 was subpar. This led to developing code in VBA that allows us to review the performance of all 12 companies we have data for over both years in order to find other investment opportunities. 

## Results
#### Stock Performance
We can see in 2017 that stocks performed well across the board for the companies we had data for. All but one had positive return over the year.

![2017_analysis](https://user-images.githubusercontent.com/90737940/135729497-0580e446-89bd-469c-9a14-49e47f321af2.png)

In 2018 performance was down with only two instances of a company achieving positive returns. 

![2018_analysis](https://user-images.githubusercontent.com/90737940/135729536-ce214c5f-ddad-4454-9f05-52d6fbfa335f.png)

Using the output from this analysis, Steve should recommend his parents invest in ENPH as there were positive returns in both years and ENPH performed well in 2018 which was a down year overall for almost all other companies in thies analysis.

#### Impact of Refactored Script
Refactoring the code allowed for a much faster run time. For this analysis we are only looking at twelve companies so the run time was not a serious issue on the original script. If we were analyzing the stock performance for thousands of companies this could be a significant issue. The refactored runtime can be seen below for 2017 and 2018 respectively:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/90737940/135729828-8d5c474a-1936-4840-9f73-4789991a0b18.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/90737940/135729832-79ee789f-ac3b-4ff1-80dd-8222edd8b0b3.png)

The key to the faster runtime is avoiding the nested loops. By eliminating the nested loop in the refactored script, runtime was increased by about 80 percent for both 2017 and 2018. In this case we saved fractions of a second. If our analysis was expanded to review thousands of companies, eliminating the nested loop has the potential to save a significant amount of time running the analysis.

By creating an array to store the output for the ticker, the ticker volume, and the return, we were able to store the outputs in the array and call on those values in the subsequent loop to be output within the worksheet. This limits the number of loops the computer has to run through compared to a nested loop. Within a nested loop the number of times the computer has to loop increases exponentially by the number of loops and the number of iterators within each loop. The nested loops in the original code which leads to the exponential growth in loops required is highlighted below:

![Original_Code_](https://user-images.githubusercontent.com/90737940/135730778-c7024eb3-f5fc-4ead-9d54-6b7adc0e9c41.jpg)

## Summary
#### What are the advantages or disadvantages to refactoring code

There are a few key benefits to refactoring code. First, by refactoring the code we can improve the runtime of a program. This can be very important when we are looking at large datasets. Another benefit to refactoring is to make the code easier to understand and help eliminate potential errors in the code. As the module mentions, refactoring can be done by another person. This can provide a good opportunity for another set of eyes to find the shortfalls of the code and to make sure it is understandable to another person.

While there are many benefits to refactoring code, there are also a couple key disadvantages. Refactoring can be time consuming. If you or someone else is refactoring your code, it might take a significant amount of time and have little performance increase over the original code. There is also an additional opportunity for errors. If a mistake is made, additional time must be spent to find and fix what went wrong.

#### How do these pros and cons apply to refactoring the original VBA script

There were definitely a couple key benefits and drawbacks to refactoring our original script. The first key benefit was the educational aspect of reworking our code. We also did speed up the runtime on the original script and make it easier to understand. That said, in a real world setting refactoring this code would be unnecessary as the original code was not slow enough to provide any real issue analyzing this amount of data. The time spent refactoring the code would not have been worth the amount of time saved running the code. If we were analyzing significantly more data this would be a different story.

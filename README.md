# Stock-Analysis
Repository for stock analysis exercise and challenge for the VBA module. Refer to All Stocks Analysis Refactored sheet in VBA_Challenge.xlsm for the challenge exercise

## Overview of Project
To enable Steve to give the best fund diversification option to his parents using the stock data provided. This involved analyzing all the green energy stock options available from past couple of years to see which among the lot provided the best return on investment (ROI) %. This would help Steve to suggest the best investment options for his parents. Data driven investment suggestion would definitely reduce the risk of losses and would allow in diversifying their investments in the best possible stocks. 

### Purpose
Considering the future of green energy companies, Steve's parents were planning to invest all their money in a company called Daqo (DQ). Steve, having a finance degree wants to do a thorough data analysis on Daqo stocks plus other stocks in this area and give his parents the best possible investment option to maximize returns.

The first step of the analysis is to understand how the DQ stocks performed in the previous year by calculating the return for DQ stocks.

The next step is for Steve to be able to perform a similar returns on investment analysis on other stocks of companies in the green energy area. 

The purpose of this project is as follows:

	- To provide Steve with a flexible code that can run on a group of stocks for different years and calculate the return on investment % for each. 
	- The output should be readily readable possibly by formatting the output in a way that helps Steve to understand the good and bad stocks by a quick glance at the results (in terms of returns on investment %)
	- The code should be well annotated to help Steve understand the purpose as well as enable him to extend it later if needed
	- The code should be written efficiently so that if needed, it can also be executed on a larger dataset


## Results

### How does the list of green energy stocks perform in 2018 compared to 2017
In 2017 all the green energy stocks in Steve's analysis list performed really well showing a very good percentage return on investment. We can see from the following analysis output that all stocks (except one - TERP) had a positive ROI %. See the number of greens in the analysis image below:

![2017 Stocks Analysis](/Resources/StockPerformance_2017.png)

In 2018, most of the stocks did not perform too well and had a negative ROI %. See analysis below and check the number of red cells in the last column:

![2018 Stocks Analysis](/Resources/StockPerformance_2018.png)

Looking at **both 2017 and 2018 stock performance** it does not make sense for Steve's parents to invest only in Daqo stocks. It's a good idea to consider other options like investing in an additional green energy stock company or even considering to invest in stocks of companies in other areas.


### How efficient is the code when used on a bigger dataset in future?
The code does provide the desired analysis but it should be written in a way so that it can handle bigger datasets in future as Steve can use this in later years for different clients. To help understand how efficiently the code executes in terms of time, the execution time is measured using a timer. With the original code, for 2017 data, the execution was completed in approximately 0.74 seconds and for 2018 data it took about 0.75 seconds. As a data analyst, a second glance at the code did indicate code segments that could be refactored to improve the code performance - removing the nested for loops would definitely reduce the execution time. With this in mind, the code was restructured and helped in significantly reducing the execution time. 

The following images give the **comparative execution times for 2017 data analysis** - the first image shows the time with the original code and the second shows the time with the refactored code which **reduced to 0.25 seconds from 0.78 seconds**.

![2017 Stocks Analysis Time](/Resources/2017_Stock_Original_Time.png)

![2017 Stocks Analysis Refactored Time](/Resources/VBA_Challenge_2017.png)


The following images give the **comparative execution times for 2018 data analysis** - the first image shows the time with the original code and the second shows the time with the refactored code which **reduced to 0.24 seconds from 0.74 seconds**.

![2018 Stocks Analysis Time](/Resources/2018_Stock_Original_Time.png)

![2018 Stocks Analysis Refactored Time](/Resources/VBA_Challenge_2018.png)
 
## Summary

### What is code refactoring?
Code Refactoring is the process of restructuring the code without adding any new functionality. The purpose of refactoring is to make the code more efficient — by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read, understand and enhance the code. 

### Advantages of code refactoring
- Increases code readability and maintainability 
- Makes the code more understandable
- Helps write efficient code in terms of shorter execution time and less memory usage
- Helps design the software better in the long run
- Reduces the incidence of bug introduction in the code 
- Significantly improves code reusability

#### Advantages of code refactoring in the current context
- The refactored code is more efficient in terms of execution time and memory usage 
- The code is now better suited to be used on future analysis involving larger datasets
- The code is well annotated and more readable making it easier for future developers or end users to understand and enhance the same if needed
- The code has better logic and structure by getting rid of the nested for loop


### Disadvantages of code refactoring
- Code refactoring can be risky for bigger code base where the overall functionality is difficult to understand
- It can introduce defects specially in cases where there are no well defined test cases for the functionality
- It can be time consuming and may push the deadlines if not planned properly

#### Disadvantages of code refactoring in the current context
- Using of different arrays in the refactored code can be difficult to comprehend for a  beginner developer 
- If the refactored code does not have proper comments for each step then it may not be very understandable as it looks quite different from the original code







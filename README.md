# Analysis of Green Energy Stocks Utilizing Excel Macros

Analyzed 12 green stocks, utilizing 2017 and 2018 data, in order to understand trading activity and yearly returns.

## Purpose

This project was intended to provide insight on appropriate green stocks in which to invest.  This project was conducted utilizing excel, and creating a macro in order to automate the analysis.  The initial analysis centered around the stock "DAQO" to determine the total daily volume and return during 2018.  This was performed utilizing the ticker symbol, date, daily closing price, and daily volume.  The  analysis was expanded to include 11 other green stock over 2017 and 2018.  Finally, the macro was refracted in order to run more efficiently to expand the dataset to include the entire stock market over the last few years.

## Results

The data from the analysis:

![AllStocks2017](https://user-images.githubusercontent.com/85590155/123531328-de811b00-d6c0-11eb-9393-928260b34eee.PNG)

![AllStocks2018](https://user-images.githubusercontent.com/85590155/123531332-e93bb000-d6c0-11eb-9087-a313150585df.PNG)

The stock performance of these 12 companies were stronger in 2017 than 2018.  Further investigation is needed to determine if this industry is an attractive investment option.  For instance, do these stocks track the performance of the broader market, or were there macro trends within the industry that led to this decline in performance.  However, we can see that both "ENPH" and "RUN"U delivered strong returns for 2018.  Also, both were actively traded which gives confidence that the price accurately reflects the value of the company.  

## Analysis

Since the next step is to perform an analysis of the entire market, the VBA code the supports the macro was refracted in order to reduce the processing time.  The original code was:

![Original Code](https://user-images.githubusercontent.com/85590155/123531667-c3fc7100-d6c3-11eb-85a1-48f2d919cb99.PNG)

This code looped through the entire data set for each of the tickers.  The VBA code was refracted to:

![Refracted Code](https://user-images.githubusercontent.com/85590155/123531695-00c86800-d6c4-11eb-8e58-8a72b017f49a.PNG)

This code only loops through the data once.  As we can see, this dramatically decreases processing time:

2017 Original Code

![2017performance](https://user-images.githubusercontent.com/85590155/123531755-6288d200-d6c4-11eb-9d3d-4c9384af5bc1.PNG)

------

2017 Refracted Code

![VBA_Challenge_2017](https://user-images.githubusercontent.com/85590155/123531761-72081b00-d6c4-11eb-9bff-c9c145b7fecc.PNG)

-----

2018 Original Code

![2018performance](https://user-images.githubusercontent.com/85590155/123531764-7c2a1980-d6c4-11eb-8c5e-48c2ea4c9bf3.PNG)

-----

2018 Refracted Code

![VBA_Challenge_2018](https://user-images.githubusercontent.com/85590155/123531769-877d4500-d6c4-11eb-9bd6-4131a900ac25.PNG)

In both cases, it is over 70% increase in performance.  While the time savings is not significant for 12 stocks, it can be quite impactful performing an analysis over thousands of stocks.

## Summary

While the initial analysis provides useful information, more detailed analysis is necessary.  The output of this project provides a macro which will automate that analysis.  Work was performed to streamline the VBA code in order to achieve optimal performance.  The time savings when dealing with large data sets can be substantial.  But refactoring does have it's drawbacks.  The time spent refactoring might not be worth the gain depending on how the code will be utilized.  Also, if other people will be working with the code, the simplest solution might be the better path rather than the optimal solution.  But in this case, considering more analysis on the broader market will occur, we are confident that the refracted macro is the appropriate solution.

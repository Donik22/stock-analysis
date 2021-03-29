# stock-analysis
Analyze to choose the best alternative energy Stock.

## Overview of Project

In this project, we will be using VBA and Excel to analyze thousands of stocks. We are taking into account total daily Volume and Annual Return to determine which stock to go for.
We Also used the code to format the table in such a way that Steve(Client) Can get a Clear idea with just a glance at the results.

### Purpose

The purpose of this Analysis is to help steves Parents To Pick the best Sustainable Energy stock. Predicting the future return on investment of each ticker by taking a look at the data from 2017 and 2018 Data. The code should run through every ticker in every row. We will also be refactoring our code to find the fastest and most efficient way to run it.

## Analysis And Challenges
### Analysis Base On 2017 Data

The Year 2017 Was a Good Year For green Stocks Except For **TerraForm Power (Ticker: TEPR)**. We Noticed stocks such as **Daqo New Energy(Ticker: DQ)** go as high as 199.4% in annual return. As shown below All but one Stock Increased in value.

![Table_2017](https://github.com/Donik22/stock-analysis/blob/main/Resources/All_Stocks_2017_Table.PNG)

### Analysis Based on 2018 Data

The year 2018 Witnessed a drop in almost all green stocks some more than others but overall it was a bad year for green stocks. we notice that two stocks maintained growth during these two year and they are **Enphae Energy(Ticker: ENPH)** & **Sunrun(Ticker: SUN)**. Referring to the attached picture the two stocks had two of the highest volume traded.

![Table_2018](https://github.com/Donik22/stock-analysis/blob/main/Resources/All_Stocks_2018_Table.PNG)

### Challenges and Difficulties Encountered

Challenges Included Enabling macros, Designating appropriate Variables/Arrays, Formatting, And Refactoring our data to Run Faster and more efficiently. The Unrefactored Script had a runtime of **0.65 Seconds** Upon Refactoring We got the code to do more in less time. Our Refactored Script Runtime was **0.14 Seconds** Which is 4.6 times Faster than the Initial Script. Images Of the refactored runtime area attached below.

![Table_2018](https://github.com/Donik22/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)


![Table_2018](https://github.com/Donik22/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

## Conclusion
Upon Analyzing We advise steves Parents To look into Two Stocks Namely **Enphae Energy(Ticker: ENPH)** & **Sunrun(Ticker: SUN)**. Those two stocks are the most traded and had an Excellent Two Year Run. Avoid Stocks like **TerraForm Power (Ticker: TEPR)** and **Daqo New Energy(Ticker: DQ)** due to negative returns and high volatility. The criteria for picking Or Avoiding these stocks are stable increase and Volume Traded. The Script used to Produce these outputs will be attached below.

[VBA_Script](https://github.com/Donik22/stock-analysis/blob/main/VBA_Challenge.vbs)

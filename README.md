# Stock Analysis Module 2 Challenge

## Overview of Project:
In this project, our client Steve asked us to prepare a workbook for him to analyze Stock Market data to determine each stock's Total Daily Volume and their Return %. This will help Steve to determine the right stocks to recommend to his parents that are looking to invest. We took the inital data set provided in Excel and used VBA to create an efficient, user-friendly workbook for our client to be able to easily click a button and enter the year and then have the workbook quickly display the results. 

## Results
### **Stock Performance 2017 vs 2018**
When analyzing the performance of the stocks in 2017 versus 2018, there is a significant change in the returns on the stocks in 2018 than in 2017. In 2017, all except one of the stocks (TERP) has a positive return for the year in terms of Total Daily Volume. But in 2018, only two stocks had a positive return (ENPH and RUN). The remaining 10 stocks had a negative return. 

#### **2017 Results**
![2017 Results](./Resources/VBA_Challenge_2017_Results.png)

#### **2018 Results** 
![2018 Results](./Resources/VBA_Challenge_2018_Results.png)

### **VBA Code Analysis**
![Setting Variables](./Resources/VBA_Challenge_Setting%20Variables.png)
In order to create the analysis tables in Excel to provide to Steve, our client, we started off setting our variables to allow us to later loop through the data in order to find the Total Daily Volume and ultimately the Return %. 
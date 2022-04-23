# Stock Analysis using VBA (Microsoft Excel)

 
## **Overview**

The Stocks Analysis project looks at performance of companies based on their stocks in the past couple of years. I summed up the total volume of each company and calculated the percent difference of their starting and ending price. The purpose of the project was to refactor the code we built in Module 2 to be able to analyze stocks data in a faster way.

You can find the Excel Spreadsheet with the VBA scrip here [VBA Challenge](https://duckduckgo.com).

## **Results**

**Stock performance:** <br>
In 2017, SPWR had the highest total volume with 782M and FSLR having the second highest total volume with 684M. The bottom performers in terms of volume were HASI and DQ with volumes of 35M and 82M, relatively. In 2018, ENPH and SPWR having the highest total volume with 607M and 538M, respectively. When comparing between the two years, the total volume is more spread out across the companies in 2018 than in 2017. 

When looking at performance based on returns, in 2017, majority of the stocks saw positive returns with only TERP seeing a negative return of -7.2%. The market did take a turn in the next year after for most companies saw negative returns with only two, ENPH and RUN. 


**Runtime performance:** <br>
Refactoring the code decreased the runtime of the program. The first iteration of the code had a runtime of more than 1 second but with this refactor, the runtime was cut down to ~0.2 seconds (the images below show the runtime of the program coming out the refactor). <br>
![Runtime for 2017](Resources/VBA_Challenge_2017.png)
![Runtime for 2018](Resources/VBA_Challenge_2018.png)

The refactor had the output variables as arrays and a variable to use to index the counter's position, which made the program run faster. The code below is the example where I created the output arrays. 
> '1a) Create a ticker Index
>    tickerIndex = 0 <br>
> '1b) Create three output arrays <br>
>   Dim tickerVolumes(12) As Long <br>
>   Dim tickerStartingPrices(12) As Single <br>
>   Dim tickerEndingPrices(12) As Single <br>


## **Summary**

**Original VBA vs. Refactored VBA** <br>
The original VBA script relied on a nested for-loop, whereas the refactored VBA did not. In addition, the original VBA script looked to variable outputs whereas the refactored VBA used arrays as variables. These differences between the two scripts made a significant performance change in the runtime. The refactored script was significantly faster than the original script, the run time decreased from around 1 second to run to around 0.2 seconds.

**Advantages of the Refactor** <br>
The refactor to the code was pretty much a restructure of the code. The restructure made the code cleaner, which is an advantage which made the code easier to follow. The refactor also allowed for the code to be more optimized. Having changed the code to be more optimal decreased the runtime of the program.

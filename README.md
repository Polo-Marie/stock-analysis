# stock-analysis

## Overview of Project:
The purpose of this analysis was to refactor the VBA code that collects information on stock market data for years 2017 and 2018. Code has already been created, but it is not an efficient process due to the amount of data collected on the stocks, so this challenge is being done to produce a more efficent and quicker process of getting through the data. By refactoring the code, our code can end up taking fewer steps, using less memory, or even lead to improved logic in its processing. It often lead to quicker processing and even cleaner code for reading and following. 

Refactoring is a common process, and here on this project, we attempt the refactoring on a new workbook to lead to a quicker process. This will be demonstrated through the use of a timer and comparing our original code with our new refactored code to show the improvement of efficiency in our code's process.

## Results
### Original Code

The timer was run on the original VBA code and the following time was given for processing:

![](Resources/2017_Original_VBA.png)

![](Resources/2018_Original_VBA.png)

It can be seen that in the original VBA coding, it takes 7103.10 seconds and 7071.36 seconds for code processing for years 2017 and 2018, respectively. The VBA codes are written to consider all 12 tickers within the given file. Those tickers are listed as their own array and are individually considered within the code process. This takes time to consider each and every ticker against the array. 

To improve the length of time to process the VBA code so that it will be capable of handling a high number and volume of tickers, instead of processing the code against every ticker, we have refactored the code to loop through rows and process itself against the row above and below it, this way, it efficiently can decipher the different tickers from each row vs being highly dependent on processing each ticker of each row against the array against potentially hundreds, if not thousands, of stock tickers.

Above, we have already provided the timing results of the original VBA code. The following photos contain the new timing of the refactored VBA code:

![](Resources/2017_Refactored_VBA.png)

![](Resources/2018_Refactored_VBA.png)

It can be seen from these photos that with the refactoring, the timing of code processing has significantly decreased:
  * 2017's VBA code changed from 7103.10 seconds to .1811 seconds
  * 2018's VBA code changed from 7071.36 seconds to .1650 seconds


## Conclusion
To conclude this analysis, let's ask and then answer the following questions:

  1. What are the advantages or disadvantages of refactoring code?
  2. How do these pros and cons apply to refactoring the original VBA script?

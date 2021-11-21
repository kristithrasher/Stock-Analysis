# Stock Analysis

## Background
Steve first came to us fresh out of college with a finance degree, His parents were his first clients. They were interested in investing all their money in DQ stock per is parents’ request. After his research he felt it was in his best interest to look at more stock options for his parents, so he wanted some help in accessing is data more efficiently. Initially Steve gave us an excel file of twelve different stocks for two years. 
We agreed to help Steve and presented him with a workbook that ran analysis on the data set he gave us with the 12 stocks and created a button that would automate the code we wrote to analyze and output the results to a new worksheet we titled “All Stocks Analysis”. We had simple user-friendly buttons created to automate clearing the worksheet, running analysis and a user prompt to insert the year he wanted to run the data on. This automation made it easy for Steve to enter the date he wanted to run his analysis on. We also created a macro that would send a message box to the screen with code performance results including how long it took the analysis code to run and output the data. The code is highlighted by positive or negative returns by color. Green represents the positive returns and red for negative returns. This makes it easy for Steve to read more easily. 
After presenting Steve with the workbook, he was pleased. He wanted to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might not work as well for thousands of stocks. The existing code would take long time to execute a larger dataset so we refactored our current code to make it more efficient by taking fewer steps and using less memory.  

## Results

### The Results of the initial set of code

The initial code that we used required twelve separated iterations over the year inputted by user of stock data for each in stock of interest. It had to run through the data of the worksheet twelve times and produced us with the results we wanted. 

The results of the original code for each of the two years analyzed is shown below 
  
<img width="230" alt="2021-11-20 (24)" src="https://user-images.githubusercontent.com/94208810/142775008-0aa87ccc-8124-427e-a13e-32cfe23eb505.png">
<img width="230" alt="2021-11-20 (25)" src="https://user-images.githubusercontent.com/94208810/142775017-6f5c13e8-014f-4fe5-bc1e-328fe8005e91.png">


### The results of the refactored code

The refactored code iterated over all the data for the year the user specified only once resulting in same output of all data needed to output.
Creating a tickerIndex variable that is set to 0 before iterating over all the rows of data. This variable accessed the correct index across the four different arrays (ticker, tickerVolumes, tickerStartingPrices and tickerEndingPrices.  We increased the current tickerVolumes by using the tickerIndex variable as the index, with following code and declared array variables.  
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex + cells(i, 8).value
This allowed our code to be more efficient and run faster. 

The results show it to be more efficient in time. The results each year is shown below. 
  <img width="234" alt="2021-11-20 (21)" src="https://user-images.githubusercontent.com/94208810/142775031-2f4abe7b-b485-4b2b-9866-5747d903f1ea.png">
<img width="234" alt="2021-11-20 (22)" src="https://user-images.githubusercontent.com/94208810/142775033-b4f5360d-0299-4887-b64a-a50e1a85c2ba.png">

## Summary

### The Advantages and Disadvantages of refactoring code. 
Advantages to refactoring code is you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

### The Disadvantages of refactoring code. 
Disadvantages of refactoring code is it cannot fix underlying architecture problems or change the functionality of the code. It also can be risky in creating error especially if a big application or programmer doesn’t fully understand what it is they are improving. 

### The advantages and disadvantages of the refactored VBA script and the original code are 
The refactored code ran faster than the initial set of code. It also allows for it to gather more data efficiently by creating/defining new variables and declaring arrays. We have one big for loop to do this instead of iterating through the same data twelve seperate times. 


### Challenges and Difficulties Encountered
I had trouble at times keeping the data organized from the old code vs the new code.


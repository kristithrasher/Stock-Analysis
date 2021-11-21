Initial 
# Stock-Analysis

## Background
Steve is our client, he just graduated with a finance degree. Stevea came to us and asked us to help him analyze some stock data. He is investing his parents money as they are his first clients. His parents want to invest in Green Energy Stock and have decided they wanted to invest all into DAQO stock or "DQ". This is a green energy company that makes silicone wafers for solar panels.  Steve is concerned aoubt diversifying their option. He has given us an excel file with twelve different stocks including DQ for two different years, either 2017 or 2018.  in the two worsheets were the tickers for each Stock, date the stock opened, open price, high price, low price,close price, adjust close price, and volume.  Steve realizes he will need to enable macros if he wants to run the analysis on in the future. We will help Steve by creating Macros using Visual Basic Application and changing the file type from .xlsx to .xlsm (macro enabled). 

Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.






## Results

  Analysis is well described with screenshots and code.
Steve realizes he will need to enable macros if he wants to run the analysis on new data in the future. We have enabled macros so Steve can run the analysis whenever he wants. 
We began by renaming the Excel file to VBA_Challenge.xlsm. This .xlsm file extension extension are files that contain macro-enabled spreadsheet files that have been created with the Microsoft Excel spreadsheet application. The macros are now enabled to run anytime.  We used Visual Basic Application to create macros to run analysis on multiple stocks. We created an output sheet titled "Stock Analysis" in our Excel workbook. We created a Macro called All_Stocks_Analysis to output the ticker, total daily volume and the return percentage of each stock. The total daily volume is the amount of times a stock was traded and the return is the yearly return is how much your investment grew or shrunk by the end of the year. to determine how a stock perfomrs for that year is by finding the yearly return. yearly return =  the percentage increase or decrease in price from the beginning of the year to the end of the year. We also created a macro to automatically format the data. We highlighed positive returns in green and negative in red.
The table is easy to read, We Highlighted by color green to determine at a glance which stocks performed well and highlighted in red which ones did not.
We created a button on our sheet to make it more user friendly so steve can easily run his analysis on all Stocks. Within that Macro we have it to where it runs the data on the year the user inputs. So in Steve's case he can enter 2017 or 2018 and get Analysis on each year for all Stocks. We measured code performance and wanted to display to show Steve how fast it took to run the analysis. We did this also inserted a colde block to gather the time it took to run the macro. Once the analysis is performed a pop up appears and shows the user the amount of time for the year. Steve was estatic to see how  fast the script ran. We refactored our code to run all stock analysis for

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s




## Summary
  There is a detailed statement of the advantages and disadvantages of refactoring code in general.
  
  There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script.  


### Challenges and Difficulties Encountered


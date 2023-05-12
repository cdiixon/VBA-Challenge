# VBA-Challenge in Summary/Conclusion:

Firstly all variables that are going to be run with in the script must be defined

Next create new headers for the values we want to find within each worksheet

Start by finding the last value in row A so that the code runs dynamically regardless of if the final line is the same or not across worksheets (LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row)

Then using the defined variables calculate and display the values for Yearly Change (conditional format so - values are red and + are green), Percent Change (repeat conditional formatting), and Total Volume ensuring
that each is stored in the correct cell and in the correct format

repeat the process for every x value in column A in order to ensure all tickers are aqcuired and all number calculations are accurate

Next find the last Value in the newly created row I so that once again it is dynamic incase the number of tickers in a given year is variable (LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row)

Then using our defined variables calculate which ticker had the Greatest Volume, Greatest % Increase, and Greatest % Decrease ensuring that each is stored in the correct cell and in the correct format

repeat the process for every x value in column I in order to ensure all tickers are aqcuired and all number calculations are accurate

Automatically adjust all columns so that all of our data fits properly (Worksheets(Work_Sheet).Columns("A:AB").AutoFit)

repeat the process for the next worksheet

In Conclusion we can use this data to show us which stocks showed the most growth, the most decline, and which stock was overall most successful with a given calendar year!

2018:
Greatest % Increase - THB - 141.42%
Greatest % Decrease - RKS - -90.02%
Greatest Total Volume - QKN - 1.69E+12

2019:
Greatest % Increase - RYU - 190.3%
Greatest % Decrease - RKS - -91.60%
Greatest Total Volume - ZQD - 4.37E+12

2020:
Greatest % Increase - YDI - 188.76%
Greatest % Decrease - VNG - -89.05%
Greatest Total Volume - QKN - 3.45E+12

Based on this investors may want to avoid stocks such as RKS which saw the biggest % decrease of any stock across 2 years in 2018 and 2019.
Investors could feel more confident investing in a stock such as QKN which finished with the greatest Total Volume of any stock across 2 years in 2018 and 2020

AnalyzeStocks VBA Script
Overview
The AnalyzeStocks VBA script is designed to analyze stock data across multiple worksheets in an Excel workbook. The script calculates the quarterly change, percentage change, and total volume for each stock ticker, and identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. It outputs the results in the worksheet and applies conditional formatting based on the quarterly change.

Features
Calculates the quarterly change for each stock ticker.
Computes the percentage change from the opening price to the closing price of the quarter.
Aggregates the total stock volume for each ticker.
Identifies and highlights the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume.
Applies conditional formatting to highlight positive and negative quarterly changes.
Instructions
Prerequisites
Microsoft Excel with VBA support.
Steps to Use the Script
Open the Excel Workbook:
Open the Excel workbook containing the stock data. Ensure each worksheet follows the same format with columns for Ticker, Date, Open Price, High Price, Low Price, Close Price, and Volume.

Open the VBA Editor:
Press Alt + F11 to open the VBA editor.

Insert a New Module:
Go to Insert > Module to insert a new module.

Copy and Paste the Script:
Copy the AnalyzeStocks VBA script provided below and paste it into the newly created module.

Run the Script:
Close the VBA editor. Press Alt + F8, select AnalyzeStocks from the list of macros, and click Run.You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data.

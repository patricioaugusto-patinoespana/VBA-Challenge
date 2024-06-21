# Multiple Year Data Stock

In this file you will see the resume of the Ticker Symbol and know the Quarterly Change by Ticker Symbol with the Open Price, and Close Price, 
and wit this information you will also know the Percent Change, and the total Stock Volume per Ticker Symbol. 
Also you will now the Greates % Increase, the Greates % Decrease and the Greatest Volume, for the 4th Quarter Sheets. 

# Instructions
You need to Open the file Multiple_year_stock_data and Enable Macros, in the file you will see 4 sheets with Q1, Q2, Q3, Q4, with all the data: 
- ticker
- date
- open
- high
- low
- close
- vol
You have to press the Button Year Stock Data to get all the information. See the image for the first view of the file:
<img width="1090" alt="VBA CHALLENGE INITIAL" src="https://github.com/patricioaugusto-patinoespana/VBA-Challenge/assets/139070248/ee9387e3-7663-49fd-a35b-d84eb945f613">

# File Structure
The file has 4 sheets for Quarters of the year, so we have Q1, Q2, Q3, Q4, in each quarter is divided for ticker symbol and each ticker symbol has:
- ticker
- date
- open
- high
- low
- close
- vol
So with this information we are going to get information for the 4 sheets for each Quarter, with each Ticker Symbol.
- The quarterly Change with the Close Price - Open Price by Tikcer symbol
- The Percentage change that every Ticker Symbol has
- The Total Stock Volume for each Ticker Symbol
- The Ticker Symbol with the Greatest % Increase
- The Ticker Symbol with the Greatest % Decrease
- The Ticker Symbol with the Greatest Volume

# Code Description
1.- Set Variables 
2.- Establish that the code run in all sheets
3.- Start Counters 
4.- Set headers, we need to set the new headers we are going to use
5.- Set the Last Row so the code knows were to stop, because every sheet ends in different rows
6.- Go to the entire row and find the values:
    - Identify were is the Ticker_Name, Open, Price, Close Price and Volume 
    - Set the values:
      - Each Ticker Symbol in the new column, resuming all the ticker symbols we have 
      - The Quarterly Change that is Close Price - Open Price
      - The percent Change That is Quarterly Change / Open Price 
    - Put the values in the new cells 
7.- Set color for Quarterly Change: Green - Positive , Red - Negative
8.- Establish Max Increase, Max Decrease An Max Volume and know what ticker symbol has it and the percentage or the volume 
9.- Write the results for Max Increase, Max Decrease and Max Volume in the new cells we establish for this values 

# Results 
In the first part of the code we can observe that we have the resume of every Ticker Symbol, The Quarterly Change, positive - green, negative -red,
the percentage change for each Ticker Symbol, and also the Total Stock Volume for each Ticker, as we observe in the image: 
<img width="770" alt="VBA CHALLENGE (1)" src="https://github.com/patricioaugusto-patinoespana/VBA-Challenge/assets/139070248/e272e275-5509-49bc-a1b7-0853f237b017">
In the second part we have the Ticker that has the Greates % Increase, the Greates % Decrease and the Greatest Volume as we obverse in the image: 
<img width="1087" alt="VBA CHALLENGE (2)" src="https://github.com/patricioaugusto-patinoespana/VBA-Challenge/assets/139070248/55468f0b-9c40-4ae5-b0b6-5a99750340a0">

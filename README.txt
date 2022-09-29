#  Ticker Data Filler Application

This application automatically fills in certain data given an excel file with stocks tickers and a specified date placed in the first two columns


The headers should be in this order to correctly fill data:
Ticker, Date, Gap %, Open, High, Low, Close, Volume

You may also choose to have multiple days of data from the starting day and in that case add on any number of columns in the headers just making sure each day has its own Open, High, Low, Close and Volume columns

An example of an excel sheet's headers wanting to track two days of data per ticker would be:

Ticker, Date, Gap %, Open, High, Low, Close, Volume, Day 2 Open, Day 2 High, Day 2 Low, Day 2 Close, Day 2 Volume

Make sure that headers match the requirements, otherwise no data will be filled or will be incorrectly filled

Note: remember you can only fill data for the past, so if a row consists of a ticker with today's date and you are requesting two days of data, the script cannot fill the row.

Note: Some ticker data may not fill due to errors with TD Ameritrade's API or if the stock has been delisted.


# How to use the program

1. Assuming the headers in the excel file are correct, run the executable

2. Follow the prompt asking for excel filepath and paste filepath

3. Follow the prompt asking for which sheet to edit

4. Watch the data get filled!



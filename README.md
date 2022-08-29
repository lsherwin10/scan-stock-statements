# scan-stock-statements
This project uses the Robinhood APIs to pull holdings from their servers, build a DataFrame of relevant information, and calculate more values to add to the spreadsheet. It also generates pie charts to display diversification information, and the data is exported to a spreadsheet for any future usage.

I have also added the ability to parse a CashApp monthly statement found in the CashApp mobile application. You simply provide the PDF in the project folder (not inside src) and name it `cashapp.pdf`. The code will import your holdings and add them to the final result.

The code will generate two files: a `Diversification.png` image, and an `Investments.xlsx` file. The image is also added as a worksheet inside the spreadsheet. The spreadsheet will contain a summary sheet, a sheet with all of your individual holdings, and the diversification image.

To run this code, simply clone the repository. Change your working directory in Terminal to the project directory, and execute `python3 src/driver.py`. This will require you to log in with Robinhood and 2FA, and the CashApp document will be parsed if it is present. The final output will be saved to the project directory.

You can also import this package by using `import driver` and use `driver.run(<filename>)` to run this with or without a provided file path. Not providing a filepath will simply run the Robinhood side, while providing one will run both and combine the two.
# Portfolio-Management-Tool
This project is designed to assist in portfolio management by automating buy-sell transaction tracking, valuation, and profit and loss (P&L) calculation. It processes stock data from Excel files and performs detailed financial computations including FIFO-based P&L, valuation on specific dates, CAGR, XIRR, and capital gains tax calculations.
# Note: I am attaching the link to access user side file(Buy-sell sheet).
[Buy-Sell Sheet](https://docs.google.com/spreadsheets/d/1ElRp3Ykaif59Nn04vRN4h16Kf6UDTlMs/edit?usp=sharing&ouid=104400550535866146904&rtpof=true&sd=true)


# Features<br>
-**Buy-Sell Queue Processing:** Automatically segregates buy and sell orders and applies a FIFO methodology to calculate profits and losses.<br>
-**Valuation:** Retrieves stock prices on specified dates and computes portfolio valuation.<br>
-**P&L Calculation:** Computes P&L for transactions, including percentage-based calculations.<br>
-**CAGR & XIRR Calculation:** Determines Compound Annual Growth Rate (CAGR) and Extended Internal Rate of Return (XIRR) for portfolio performance evaluation.<br>
-**Tax Reporting:** Segregates Long Term Capital Gains (LTCG) and Short Term Capital Gains (STCG) based on the holding period.<br>

# Prerequisites
- Python 3.x
- Required python libraries: pandas, openpyxl, pyxirr, yfinance
You can install dependencies using:<br>
```bash
pip install pandas openpyxl pyxirr yfinance
```
# Files
`main.py`: Contains the core logic for processing portfolio data, computing P&L, and generating reports.<br>
`Buy-Sell.xlsx`: The Excel file where buy-sell transactions are recorded.<br>
`Holdings and valuation.xlsx`: The Excel file for holding portfolio valuations, including P&L sections.

# How to Use
1. **Prepare the Excel Files:**<br>
   -`Buy-sell.xlsx`: Ensure your buy/sell transactions are recorded in the "Buy-Sell Entry" sheet with ISIN, symbol, quantity, dates, and net amount.<br>
   -`Holdings and valuation.xlsx`: Prepare the sheet "Holdings, Valuations and P&L" for input/output of holdings data, valuation dates, and P&L sections.<br>
2. **Run the Python Script:** Execute `main.py` script to process the data(Make sure to change maximum rows for each function so that all transactions can be recoreded). This will:<br>
         -Segregate the buy and sell orders.<br>
         -Process P&L and valuation data.<br>
         -Update the Excel files with calculated results (CAGR, XIRR, capital gains, etc.).<br>
3. **Check the Output:** The updated data will be saved directly into the `Holdings and valuation.xlsx` file, including:<br>
         -Processed buy-sell entries.<br>
         -Portfolio valuation based on closing prices from Yahoo Finance.<br>
         -Profit and loss calculations and performance metrics.<br>

# Function Overview:
-`load_worksheet(file_path, sheet_name)`:  Loads an Excel sheet for processing.<br>
-`buy_and_sell_q(df)`: Segregates buy and sell orders.<br>
-`process_trade_data(df)`: Prepares buy/sell data for further processing.<br>
-`formating_queues(q_buy, q_sell)`: Formats date data for easier handling.<br>
-`process_sell_buy_orders(q_sell_formatted, q_buy_formatted)`: Applies FIFO logic to buy-sell queues.<br>
-`process_and_update_sheet(Amt, Quantity, q_closing_Price)`: Combines valuation, P&L, and percentage P&L calculations.<br>
-`port_cagr(q1)`: Calculates the CAGR for the portfolio.<br>
-`calculate_xirr(df)`: Calculates the XIRR for the portfolio.<br>
-`populate_sheet`: Writes computed data back into Excel.<br>

# Screenshot of both files:

**Original Files:**
![Screenshot 2024-10-08 095017](https://github.com/user-attachments/assets/dee63900-476f-468d-ac73-2d5cd0a502c9)
![Screenshot 2024-10-08 095157](https://github.com/user-attachments/assets/71543ff5-8d80-4e4e-b421-611836ac8ac7)

**After Data Entry and running files:**
![Screenshot 2024-10-08 104649](https://github.com/user-attachments/assets/58beb61f-5461-4ceb-b896-e247bbd8c73b)
![Screenshot 2024-10-08 104627](https://github.com/user-attachments/assets/a92b2c1e-b2f4-4efa-8c1b-efdadba8c68e)


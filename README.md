# Stock Market Analysis for E-Commerce Companies

This project involves analyzing historical stock market data for major e-commerce companies such. The analysis is conducted using Microsoft Excel and Power BI to extract insights on stock price trends, trading volumes, volatility, and key performance indicators.

## Tools Used
- Microsoft Excel (Get & Transform Data, Formulas, Data Analysis Toolpak)
- Microsoft Power BI (Data modeling, DAX, Interactive visualizations)

## Project Overview

### Excel Process
1. **Data Import**: Use Excel's "Get & Transform Data" (Power Query) feature to import historical stock data from web or CSV files.
2. **Data Cleaning**: Format date columns, remove null values, and create calculated columns such as daily returns.
3. **Analysis**: Use Excel formulas for calculations, including:
   - Daily Return Formula:
     ```
     =(C2 - C1) / C1
     ```
     where Column C contains closing prices.
4. **Visualization**: Construct line charts for stock prices, bar charts for volumes, and use the Data Analysis Toolpak for regression and correlation analysis.
5. **Metrics Calculation**: Calculate KPIs like volatility, average price, maximum and minimum closing prices.
### Codes/Formulas for Analysis
In Excel
Total Traded Value:

Add a new column: = [@Open] * [@Volume].

Moving Average (50-day):

In a new column, for row 51: =AVERAGE(C2:C51) (if Close is column C). .

Daily Return:

In a new column, starting from row 3: =(C3-C2)/C2 (if Close is column C).

### Power BI Process
1. **Data Import**: Load cleaned stock data from Excel files or direct web sources into Power BI.
2. **Data Modeling**: Establish relationships and create calculated columns/measures with DAX.
3. **Key Metrics with DAX**: Example daily return measure:
4. **Visualization**: Build interactive line charts, volume bar charts, date slicers, and KPI cards.
5. **Dashboard Creation**: Combine visualizations and metrics into comprehensive dashboards for actionable insights.

###Total Traded Value:

1. Go to Data view > New column
TotalTraded = [Open] * [Volume]
2. Moving Average (20-day):
MA50 = CALCULATE(AVERAGE([Close]), DATESINPERIOD('stk_shop'[01-07-2025], 'stk_shop'[20-07-2025], -20, DAY))
3. Daily Return:
Return = ([Close] - CALCULATE(LASTNONBLANK([Close], 1), FILTER('stk_shop', 'stk_shop'[20-07-2025] = EARLIER('stk_shop'[20-07-2025]) - 1))) / CALCULATE(LASTNONBLANK([Close], 1), FILTER('stk_shop', 'stk_shop'[20-07-2025] = EARLIER('stk_shop'[v]) - 1))
4. Automated Data Import
Use Excel’s STOCKHISTORY function for historical data:
5. Relative Strength Index (RSI):

Calculate daily chang

Separate gains/losses, then use:

RSI = 100 - (100 / (1 + (Average Gain / Average Loss)))


## Results
- Interactive dashboards revealing stock price movements and volume trends.
- Key performance indicators that support market trend analysis and investment decision making.
- Identification of volatility periods through daily return calculations.

## Conclusion
This project demonstrates how stock market data can be effectively analyzed using Excel and Power BI without requiring extensive programming. The combined use of Excel formulas and Power BI's DAX functions enables comprehensive financial analysis and visualization.

---




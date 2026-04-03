# Portfolio Report Generator

Portfolio Report Generator is a program that reads investment transactions and generates a browser-based portfolio report. It gives you a consolidated view of current holdings, realized and unrealized return, dividends, detailed sale-trade history across both stocks and funds, and visual charts for total return and market value allocation.

Request access to the web app by email: [link](https://indsetsportfolioreport.web.app)

The web app includes:
- Email/password login
- A menu for uploading files and generating reports
- Per-user storage of uploaded transaction files and the latest generated report

You can also run the program locally on your own machine.

## Key features
- Reads all supported `.csv` files from `transaction_files/`
- Supports comma, semicolon, and tab-separated exports
- Handles UTF-8 and UTF-16 encoded files
- Tracks buys, sells, and dividends across stocks and funds
- Calculates FIFO-based realized gain/loss and return percentages
- Resolves ticker, exchange, and company name using Yahoo Finance data
- Includes visual charts in the report:
	- Total Return bar chart
	- Market Value allocation charts (holdings allocation + asset mix)
- Includes a Realized Overview table (combined realized sales per security)
- Includes Sale Trades tables (every sell transaction per security)

## Quick Start
1. Export transaction history from your brokers/banks as CSV files.
2. Put the files in `transaction_files/` (Files with `example` in the filename are ignored).
3. Compile and run:

```bash
javac -d out src/*.java src/*/*.java
java -cp out PortfolioReportGenerator
```

The program generates `portfolio-report.html`.

## Example Data
- `transaction_files/transactions_example.csv` contains realistic sample transactions with both gains and losses.

## Portfolio Example
![Portfolio report example](portfolio_report_example.png)

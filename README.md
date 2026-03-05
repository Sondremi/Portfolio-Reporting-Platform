# Portfolio Performance Tracker

A Java program that reads investment transactions from CSV files and generates a browser-based portfolio report.
Shows current holdings, combined realized sales, and detailed trade history per security (stocks and funds).

Useful when you want to combine portfolios from multiple banks or brokers in one overview.

## Key Features
- Reads all supported `.csv` files from `transaction_files/`
- Supports comma, semicolon, and tab-separated exports
- Handles UTF-8 and UTF-16 encoded files
- Tracks buys, sells, and dividends across stocks and funds
- Calculates FIFO-based realized gain/loss and return percentages
- Resolves ticker, exchange, and company name using Yahoo Finance data
- Includes a Realized Overview table (combined realized sales per security)
- Includes Sale Trades tables (every sell transaction per security)
- Report can be exported to PDF from the browser

## Quick Start
1. Export transaction history from your broker/bank as CSV files.
2. Put the files in `transaction_files/`.
3. Compile and run:

```bash
javac src/*.java
java -cp src PortfolioTracker
```

The program generates `portfolio-report.html`.

## CSV Notes
- All `.csv` files in `transaction_files/` are loaded.
- Files with `example` in the filename are ignored by default.
- Required columns in supported exports are `Verdipapir` and `Transaksjonstype`.

## Example Data
- `transaction_files/transactions_example.csv` contains realistic sample transactions with both gains and losses.
- To run only the sample dataset, temporarily rename the file to something like `transactions_demo.csv`.

## Portfolio PDF Example
[Open PDF](portfolio_report_example.pdf)

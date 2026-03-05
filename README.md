# Portfolio Performance Tracker

A Java program that analyzes your investment transactions and calculates comprehensive performance metrics across all your stocks and funds.

## Key Features
- Calculates both realized and unrealized returns
- Tracks performance across multiple transactions
- Handles dividends, buys, and sells
- Generates a Numbers-ready CSV with automatic stock price lookups

## How to Use
1. Export transaction history from one or more banks as CSV files
2. Place all transaction CSV files inside `transaction_files/`
3. Run `PortfolioTracker` to generate `portfolio.csv` with a combined overview

Example run:
```bash
javac src/*.java
java -cp src PortfolioTracker
```

## Example on portfolio
![](portfolio_example.png)

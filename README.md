# Poet Strategy Analyzer

A Python-based analytical tool for evaluating algorithmic trading strategies. This tool processes trade logs, calculates advanced performance metrics (Sharpe, Sortino, CAGR), and generates comprehensive Excel and PDF factsheets.

## Features
- **Portfolio Blending:** Dynamically rebalances multiple strategies based on allocated capital.
- **Performance Metrics:** Calculates CAGR, Sharpe Ratio, Sortino Ratio, Max Drawdown, and more.
- **Visual Reporting:** Generates equity curves, drawdown charts, and monthly return heatmaps.
- **Benchmark Comparison:** Downloads SPY data via `yfinance` to compare strategy performance against the market.

## Setup

1. Install the required libraries:
   ```bash
   pip install -r requirements.txt

2. Place your trade list CSV files in the input/ folder.
Files should contain standard trade log columns (e.g., 'Type', 'Date/Time', 'Net P&L %').

3. Configure your portfolio in Poetstrategies.csv

## Usage
Run the main script:
    ```bash
    python poetv1.py

Results (PDF and Excel dashboards) will be generated in the output/ directory.
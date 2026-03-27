# Stock & Crypto Portfolio Tracker

A personal finance portfolio tracker built with Python and Excel. Tracks stocks, ETFs, options, and cryptocurrency holdings across multiple brokerage accounts with live price refreshing and automated dashboard updates.

## Features

- **Multi-Asset Support** — Track stocks, ETFs, options, and crypto (BTC, ETH, SOL, DOGE, and more) in one place
- **Live Price Refresh** — Pulls real-time prices from Yahoo Finance (stocks/ETFs) and CoinGecko (crypto)
- **Excel Dashboard** — Portfolio summary with allocation breakdowns by account and asset type, plus pie chart visualizations
- **Trade Log** — Built-in trade logging sheet with automatic cost calculations
- **Automated Scheduling** — Windows Task Scheduler integration for hourly and on-login price updates
- **Multi-Account** — Supports tracking across Robinhood, Fidelity, and other brokerages

## Project Structure

```
├── Finance Project.xlsx    # Main portfolio workbook (Dashboard, Holdings, Trade Log)
├── build_portfolio.py      # Generates the Excel workbook with formatted sheets and formulas
├── refresh_prices.py       # Fetches live prices from CoinGecko & Yahoo Finance, updates Excel
├── update_charts.py        # Rebuilds pie charts on the Dashboard sheet
├── setup_scheduler.bat     # Creates Windows scheduled tasks for automatic price refresh
├── Open Portfolio.vbs      # Refreshes prices then opens the workbook in one click
└── refresh_prices.vbs      # Silent background price refresh script
```

## Getting Started

### Prerequisites

- Python 3.10+
- Windows (for VBS/BAT automation scripts)

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/GeorgeJohnson04/stock-portfolio-tracker.git
   cd stock-portfolio-tracker
   ```

2. Install dependencies:
   ```bash
   pip install openpyxl requests
   ```

3. Build the portfolio workbook:
   ```bash
   python build_portfolio.py
   ```

4. (Optional) Set up automatic price refresh:
   ```bash
   setup_scheduler.bat
   ```

## Usage

### Manual Price Refresh
```bash
python refresh_prices.py
```

### Quick Open (refreshes prices, then opens Excel)
Double-click `Open Portfolio.vbs`

### Adding Holdings
Open `Finance Project.xlsx` and add rows to the **Holdings** sheet:
- **Account**: Robinhood, Fidelity, etc.
- **Asset Type**: Crypto, Stock, ETF, or Options
- **Ticker**: Standard ticker symbol (BTC, ETH, AAPL, SPY, etc.)
- Fill in Quantity and Avg Cost — formulas handle the rest

### Updating Charts
```bash
python update_charts.py
```

## Supported Cryptocurrencies

BTC, ETH, SOL, DOGE, ADA, XRP, DOT, AVAX, MATIC, LINK, SHIB, LTC — additional coins can be added in `refresh_prices.py`.

## Data Sources

- **Stocks & ETFs**: [Yahoo Finance](https://finance.yahoo.com/)
- **Cryptocurrency**: [CoinGecko API](https://www.coingecko.com/en/api)

## License

MIT

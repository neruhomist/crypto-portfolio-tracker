# Alpha — Crypto Portfolio Tracker

Automated crypto portfolio tracker built with Google Sheets + Google Apps Script.

## Features

- Daily automatic price updates at 9:00 AM (OKX → BingX → Bybit, public APIs)
- Pulls trade history from OKX and BingX, calculates monthly purchase totals
- Automatically creates a new column when a new month starts
- Automatically adds a new row when a new token appears in trade history
- Calculates Qty (holdings) directly from exchange balances (OKX + BingX)
- Pulls USDT balance separately per exchange
- P/L (profit/loss) calculated via spreadsheet formulas

## Exchanges

| Exchange | Prices | Trades/Balances |
|---|---|---|
| OKX | ✅ | ✅ fully automated |
| BingX | ✅ | ✅ fully automated |
| Bybit | ✅ (public) | ❌ blocked by geo-restriction (Google Cloud IP) |
| Binance | ❌ blocked by geo-restriction | ❌ blocked by geo-restriction |

## Setup

1. Create read-only API keys on OKX and BingX
2. Save them in Script Properties: `OKX_API_KEY`, `OKX_SECRET_KEY`, `OKX_PASSPHRASE`, `BINGX_API_KEY`, `BINGX_SECRET_KEY`
3. Sheet structure ("investments" tab): column A — coin tickers, followed by monthly columns, `Current Price`, `Qty`,

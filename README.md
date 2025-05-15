# MT5 Historical Data Downloader

![Python](https://img.shields.io/badge/python-3.7+-blue.svg)
![PyQt5](https://img.shields.io/badge/PyQt5-5.15+-green.svg)
![MetaTrader5](https://img.shields.io/badge/MetaTrader5-5.0+-orange.svg)

A graphical application for downloading historical market data from MetaTrader 5 (MT5) with support for multiple symbols, timeframes, and export formats. Features include candlestick chart visualization, symbol presets, and bilingual English/Farsi interface.

## Features

- Download historical data for multiple symbols simultaneously
- Support for all standard MT5 timeframes (M1 to MN1)
- Export data to Excel (XLSX) or CSV formats
- Interactive candlestick charts with dark/light theme support
- Bilingual interface (English/Farsi)
- Symbol management with presets and watchlist import
- Customizable date ranges (specific range or days back from today)
- Selectable columns for export
- Progress tracking during downloads

## Screenshots

![Application Screenshot](screenshots/app_screenshot.png)
*(Example screenshot - add your own screenshot file)*

## Requirements

- Python 3.7+
- MetaTrader 5 terminal installed and running
- MT5 account (even demo account works)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/mt5-data-downloader.git
   cd mt5-data-downloader

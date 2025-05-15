import sys
import os
import json
import logging
import pandas as pd
import MetaTrader5 as mt5
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QComboBox, QPushButton, QRadioButton, 
                             QButtonGroup, QDateEdit, QSpinBox, QMessageBox, QFileDialog, 
                             QProgressBar, QCheckBox, QListWidget, QListWidgetItem, QFrame, 
                             QGroupBox, QSplitter, QToolButton, QInputDialog, QDialog, 
                             QDialogButtonBox, QListWidget)
from PyQt5.QtCore import QDate, Qt, QThread, pyqtSignal, QLocale, QTranslator
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon
import pytz
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.dates as mdates
from mplfinance.original_flavor import candlestick_ohlc

# Configure logging
logging.basicConfig(
    filename='mt5_downloader.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class Translator:
    """Handles language translations for the application"""
    def __init__(self):
        self.translations = {
            'en': {
                'app_title': "MT5 Historical Data Downloader",
                'symbols_label': "Symbols (comma-separated):",
                'symbol_management': "Symbol Management",
                'load_preset': "Load symbol preset...",
                'save_preset': "Save Preset",
                'load_watchlist': "Load Watchlist",
                'settings': "Settings",
                'timeframes': "Timeframes:",
                'export_format': "Export Format:",
                'date_range': "Date Range",
                'specific_range': "Specific Date Range",
                'days_back': "Days Back from Today",
                'start_date': "Start Date:",
                'end_date': "End Date:",
                'days_label': "Days back:",
                'output': "Output",
                'output_path': "Output Path:",
                'columns_label': "Columns to include in export:",
                'plot_label': "Select Symbol to Plot:",
                'plot_btn': "Plot Selected Symbol",
                'download_btn': "Download Data",
                'stop_btn': "Stop Download",
                'browse_btn': "Browse",
                'theme_tooltip': "Toggle Theme",
                'no_data': "No data to display",
                'price_chart': "Price Chart",
                'invalid_symbol': "Invalid symbol detected between commas",
                'no_symbols': "Please enter at least one valid symbol",
                'no_timeframes': "Please select at least one timeframe",
                'invalid_range': "Start date must be before end date",
                'no_columns': "Please select at least one column to export",
                'download_starting': "Starting download...",
                'download_complete': "Download completed successfully",
                'no_data_received': "Download complete but no data received",
                'error_occurred': "Error occurred",
                'mt5_connect_error': "Failed to connect to MT5. Please ensure MT5 is running.",
                'no_watchlist': "No symbols found in Market Watch",
                'preset_name': "Enter preset name:",
                'preset_saved': "Preset '{}' saved",
                'preset_loaded': "Loaded preset '{}'",
                'watchlist_loaded': "Loaded {} symbols from watchlist",
                'confirm_exit': "A download is currently running. Are you sure you want to quit?",
                'theme_message': "Dark theme is currently the only available option",
                'select_symbols': "Select symbols to import:",
                'watchlist_title': "Select Watchlist Symbols"
            },
            'fa': {
                'app_title': "دانلودگر داده‌های تاریخی متاتریدر ۵",
                'symbols_label': "نمادها (جداشده با کاما):",
                'symbol_management': "مدیریت نمادها",
                'load_preset': "بارگیری پیش‌تنظیم نمادها...",
                'save_preset': "ذخیره پیش‌تنظیم",
                'load_watchlist': "بارگیری واچ‌لیست",
                'settings': "تنظیمات",
                'timeframes': "تایم‌فریم‌ها:",
                'export_format': "فرمت خروجی:",
                'date_range': "بازه زمانی",
                'specific_range': "بازه تاریخ مشخص",
                'days_back': "روز قبل از امروز",
                'start_date': "تاریخ شروع:",
                'end_date': "تاریخ پایان:",
                'days_label': "تعداد روز:",
                'output': "خروجی",
                'output_path': "مسیر ذخیره:",
                'columns_label': "ستون‌های مورد نظر برای export:",
                'plot_label': "انتخاب نماد برای نمایش:",
                'plot_btn': "نمایش نماد انتخاب شده",
                'download_btn': "دانلود داده‌ها",
                'stop_btn': "توقف دانلود",
                'browse_btn': "مرور",
                'theme_tooltip': "تغییر تم",
                'no_data': "داده‌ای برای نمایش وجود ندارد",
                'price_chart': "نمودار قیمت",
                'invalid_symbol': "نماد خالی بین کاماها وجود دارد",
                'no_symbols': "لطفا حداقل یک نماد معتبر وارد کنید",
                'no_timeframes': "لطفا حداقل یک تایم‌فریم انتخاب کنید",
                'invalid_range': "تاریخ شروع باید قبل از تاریخ پایان باشد",
                'no_columns': "لطفا حداقل یک ستون برای export انتخاب کنید",
                'download_starting': "در حال شروع دانلود...",
                'download_complete': "دانلود با موفقیت انجام شد",
                'no_data_received': "دانلود کامل شد اما داده‌ای دریافت نشد",
                'error_occurred': "خطا رخ داده است",
                'mt5_connect_error': "اتصال به متاتریدر ۵ ناموفق بود. لطفا از اجرا بودن متاتریدر اطمینان حاصل کنید.",
                'no_watchlist': "هیچ نمادی در واچ‌لیست یافت نشد",
                'preset_name': "نام پیش‌تنظیم را وارد کنید:",
                'preset_saved': "پیش‌تنظیم '{}' ذخیره شد",
                'preset_loaded': "پیش‌تنظیم '{}' بارگیری شد",
                'watchlist_loaded': "{} نماد از واچ‌لیست بارگیری شد",
                'confirm_exit': "دانلود در حال انجام است. آیا مطمئنید که می‌خواهید خارج شوید؟",
                'theme_message': "در حال حاضر فقط تم تیره موجود است",
                'select_symbols': "نمادهای مورد نظر را انتخاب کنید:",
                'watchlist_title': "انتخاب نمادهای واچ‌لیست"
            }
        }
        self.current_lang = 'en'
        
    def set_language(self, lang):
        """Set the current language (en/fa)"""
        if lang in self.translations:
            self.current_lang = lang
            return True
        return False
    
    def tr(self, text_key):
        """Translate text based on current language"""
        return self.translations[self.current_lang].get(text_key, text_key)

class CandlestickChart(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi, facecolor='none')
        self.ax = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)
        self.dark_mode = True
        self.translator = Translator()
        
    def plot_candles(self, data, symbol=""):
        """Plot candlestick chart from OHLC data with enhanced styling"""
        self.ax.clear()
        
        if data.empty:
            self.ax.text(0.5, 0.5, self.translator.tr('no_data'), 
                        ha='center', va='center', fontsize=12,
                        color='white' if self.dark_mode else 'black')
            self.draw()
            return
        
        # Convert dates to matplotlib format
        dates = mdates.date2num(data['Date'])
        
        # Prepare OHLC data
        ohlc = list(zip(dates, data['Open'], data['High'], data['Low'], data['Close']))
        
        # Plot candlesticks with improved colors
        candle_colors = {
            'up': '#4CAF50',  # Green
            'down': '#F44336'  # Red
        }
        candlestick_ohlc(self.ax, ohlc, width=0.6, 
                        colorup=candle_colors['up'], 
                        colordown=candle_colors['down'], 
                        alpha=0.9)
        
        # Format x-axis with better date display
        self.ax.xaxis_date()
        self.fig.autofmt_xdate()
        locator = mdates.AutoDateLocator()
        formatter = mdates.ConciseDateFormatter(locator)
        self.ax.xaxis.set_major_locator(locator)
        self.ax.xaxis.set_major_formatter(formatter)
        
        # Enhanced chart styling
        bg_color = '#2D2D2D' if self.dark_mode else '#FFFFFF'
        text_color = 'white' if self.dark_mode else 'black'
        grid_color = '#555555' if self.dark_mode else '#DDDDDD'
        
        self.ax.grid(True, linestyle='--', alpha=0.7, color=grid_color)
        self.ax.set_facecolor(bg_color)
        self.fig.patch.set_facecolor(bg_color)
        self.ax.tick_params(colors=text_color)
        self.ax.xaxis.label.set_color(text_color)
        self.ax.yaxis.label.set_color(text_color)
        self.ax.title.set_color(text_color)
        
        for spine in self.ax.spines.values():
            spine.set_color(text_color)
        
        # Set title with symbol
        title = f"{symbol} {self.translator.tr('price_chart')}" if symbol else self.translator.tr('price_chart')
        self.ax.set_title(title, pad=20)
        self.ax.set_ylabel('Price', labelpad=10)
        
        self.draw()

class DataDownloadThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(dict)  # Emits dict of dataframes with full OHLC data
    error = pyqtSignal(str)
    log_message = pyqtSignal(str, str)  # message, level

    def __init__(self, symbols, timeframes, start_dt, end_dt, output_file, export_format, selected_columns):
        super().__init__()
        self.symbols = symbols
        self.timeframes = timeframes
        self.start_dt = start_dt
        self.end_dt = end_dt
        self.output_file = output_file
        self.export_format = export_format
        self.selected_columns = selected_columns
        self._is_running = True

    def run(self):
        try:
            self.log_message.emit("Attempting to initialize MT5 connection...", "INFO")
            if not mt5.initialize():
                error_msg = f"Failed to initialize MT5 connection. Error: {mt5.last_error()}"
                self.log_message.emit(error_msg, "ERROR")
                self.error.emit(error_msg)
                return

            try:
                self.log_message.emit("Getting available symbols...", "INFO")
                available_symbols = [s.name for s in mt5.symbols_get()]
                self.log_message.emit(f"Found {len(available_symbols)} available symbols", "INFO")
                
                result_data = {}  # Stores full OHLC data for charting
                export_data = {}  # Stores data with user-selected columns for export
                
                tf_mapping = {
                    'M1': mt5.TIMEFRAME_M1, 'M5': mt5.TIMEFRAME_M5, 'M15': mt5.TIMEFRAME_M15,
                    'M30': mt5.TIMEFRAME_M30, 'H1': mt5.TIMEFRAME_H1, 'H4': mt5.TIMEFRAME_H4,
                    'D1': mt5.TIMEFRAME_D1, 'W1': mt5.TIMEFRAME_W1, 'MN1': mt5.TIMEFRAME_MN1,
                }

                total_tasks = len(self.symbols) * len(self.timeframes)
                completed_tasks = 0
                self.log_message.emit(f"Starting download of {total_tasks} symbol/timeframe combinations", "INFO")
                
                for symbol in self.symbols:
                    if not self._is_running:
                        break
                        
                    exact_symbol = next((s for s in available_symbols if s.lower() == symbol.lower()), None)
                    
                    if not exact_symbol:
                        msg = f"Symbol {symbol} not found (case mismatch)"
                        self.log_message.emit(msg, "WARNING")
                        completed_tasks += len(self.timeframes)
                        self.progress.emit(int((completed_tasks / total_tasks) * 100))
                        continue

                    for timeframe in self.timeframes:
                        if not self._is_running:
                            break
                            
                        try:
                            self.log_message.emit(f"Downloading {exact_symbol} {timeframe} data...", "INFO")
                            rates = mt5.copy_rates_range(
                                exact_symbol, 
                                tf_mapping[timeframe], 
                                self.start_dt, 
                                self.end_dt
                            )

                            if rates is None or len(rates) == 0:
                                error_msg = f"No data returned for {exact_symbol} {timeframe}: {mt5.last_error()}"
                                self.log_message.emit(error_msg, "WARNING")
                                completed_tasks += 1
                                self.progress.emit(int((completed_tasks / total_tasks) * 100))
                                continue

                            # Create full OHLC dataframe for charting
                            full_df = pd.DataFrame(rates)
                            if full_df.empty:
                                msg = f"Empty dataframe for symbol {exact_symbol} {timeframe}"
                                self.log_message.emit(msg, "WARNING")
                                completed_tasks += 1
                                self.progress.emit(int((completed_tasks / total_tasks) * 100))
                                continue

                            # Convert and rename columns
                            full_df['time'] = pd.to_datetime(full_df['time'], unit='s')
                            if full_df['time'].dt.tz is None:
                                full_df['time'] = full_df['time'].dt.tz_localize('UTC')
                                
                            full_df = full_df.rename(columns={
                                'time': 'Date', 'open': 'Open', 'high': 'High',
                                'low': 'Low', 'close': 'Close', 'tick_volume': 'Volume',
                                'spread': 'Spread', 'real_volume': 'RealVolume'
                            })

                            # Store full OHLC data for charting with timeframe in key
                            result_data[f"{exact_symbol}_{timeframe}"] = full_df[['Date', 'Open', 'High', 'Low', 'Close']]
                            
                            # Create export dataframe with user-selected columns
                            export_df = full_df.copy()
                            available_cols = [col for col in self.selected_columns if col in export_df.columns]
                            if not available_cols:
                                msg = f"No valid columns selected for {exact_symbol} {timeframe}"
                                self.log_message.emit(msg, "WARNING")
                                completed_tasks += 1
                                self.progress.emit(int((completed_tasks / total_tasks) * 100))
                                continue
                            
                            export_df = export_df[available_cols]
                            
                            # For Excel export, we'll handle all data together after the loop
                            if self.export_format == "xlsx":
                                export_data[f"{exact_symbol}_{timeframe}"] = export_df
                            elif self.output_file:
                                # Single symbol/timeframe or CSV export
                                if len(self.symbols) == 1 and len(self.timeframes) == 1 and not os.path.isdir(self.output_file):
                                    filename = self.output_file
                                else:
                                    base_dir = os.path.dirname(self.output_file) if self.output_file else "."
                                    if not os.path.exists(base_dir):
                                        os.makedirs(base_dir)
                                    filename = os.path.join(
                                        base_dir,
                                        f"{exact_symbol}_{timeframe}_"
                                        f"{self.start_dt.strftime('%Y%m%d')}_to_"
                                        f"{self.end_dt.strftime('%Y%m%d')}.{self.export_format}"
                                    )
                                
                                try:
                                    if self.export_format == "xlsx":
                                        if 'Date' in export_df.columns:
                                            export_df['Date'] = export_df['Date'].dt.tz_localize(None)
                                        export_df.to_excel(filename, index=False, engine='openpyxl')
                                    else:
                                        export_df.to_csv(filename, index=False)
                                    self.log_message.emit(f"Successfully saved {len(export_df)} rows to {filename}", "INFO")
                                except Exception as e:
                                    error_msg = f"Failed to save file: {e}"
                                    self.log_message.emit(error_msg, "ERROR")
                                    completed_tasks += 1
                                    self.progress.emit(int((completed_tasks / total_tasks) * 100))
                                    continue

                            completed_tasks += 1
                            progress = int((completed_tasks / total_tasks) * 100)
                            self.progress.emit(progress)
                            self.log_message.emit(f"Progress: {progress}%", "INFO")
                            self.msleep(100)  # Small delay to allow UI updates

                        except Exception as e:
                            error_msg = f"Error processing {exact_symbol} {timeframe}: {str(e)}"
                            self.log_message.emit(error_msg, "ERROR")
                            completed_tasks += 1
                            self.progress.emit(int((completed_tasks / total_tasks) * 100))
                            continue

                # Handle multi-symbol/multi-timeframe Excel export
                if self._is_running and self.export_format == "xlsx" and len(export_data) > 0 and self.output_file:
                    try:
                        self.log_message.emit("Preparing Excel workbook...", "INFO")
                        wb = Workbook()
                        if len(wb.sheetnames) > 0:
                            wb.remove(wb[wb.sheetnames[0]])
                        
                        for sheet_name, df in export_data.items():
                            if 'Date' in df.columns:
                                df['Date'] = df['Date'].dt.tz_localize(None)
                            
                            ws = wb.create_sheet(title=sheet_name[:31])  # Excel sheet name limit
                            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
                                for c_idx, value in enumerate(row):
                                    ws.cell(row=r_idx+1, column=c_idx+1, value=value)
                        
                        wb.save(self.output_file)
                        self.log_message.emit(f"Successfully saved {len(export_data)} sheets to {self.output_file}", "INFO")
                    except Exception as e:
                        error_msg = f"Failed to save Excel file: {str(e)}"
                        self.log_message.emit(error_msg, "ERROR")
                        self.error.emit(error_msg)

                if self._is_running:
                    self.log_message.emit("Download completed successfully", "INFO")
                    self.finished.emit(result_data)
                
            except Exception as e:
                error_msg = f"Error during download: {str(e)}"
                self.log_message.emit(error_msg, "ERROR")
                self.error.emit(error_msg)
                
            finally:
                self.log_message.emit("Shutting down MT5 connection", "INFO")
                mt5.shutdown()
                
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            self.log_message.emit(error_msg, "ERROR")
            if mt5.initialized():
                mt5.shutdown()
            self.error.emit(error_msg)

    def stop(self):
        self._is_running = False

class MT5DataDownloader(QMainWindow):
    def __init__(self):
        super().__init__()
        self.translator = Translator()
        self.setWindowTitle(self.translator.tr('app_title'))
        self.setMinimumSize(1200, 700)
        self.dark_mode = True
        self.chart_data = {}
        self.current_chart_symbol = None
        self.download_thread = None  # Initialize download_thread as None
        self.setup_ui()
        self.apply_dark_theme()
        self.chart.dark_mode = self.dark_mode
        self.chart.translator = self.translator

    def setup_ui(self):
        # Main splitter (horizontal)
        main_splitter = QSplitter(Qt.Horizontal)
        
        # Left panel (controls)
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(15, 15, 15, 15)
        left_layout.setSpacing(15)
        
        # Header with theme and language toggle
        header_layout = QHBoxLayout()
        
        header = QLabel(self.translator.tr('app_title'))
        header_font = QFont("Segoe UI", 14, QFont.Bold)
        header.setFont(header_font)
        header.setStyleSheet("color: white;")
        header_layout.addWidget(header)
        
        # Language toggle button
        self.lang_btn = QPushButton("EN/FA")
        self.lang_btn.clicked.connect(self.toggle_language)
        header_layout.addWidget(self.lang_btn)
        
        # Theme toggle button
        self.theme_btn = QToolButton()
        self.theme_btn.setCheckable(True)
        self.theme_btn.setChecked(True)
        self.theme_btn.setIcon(QIcon.fromTheme('color-management'))
        self.theme_btn.setToolTip(self.translator.tr('theme_tooltip'))
        self.theme_btn.setStyleSheet("""
            QToolButton {
                background-color: transparent;
                border: none;
                padding: 5px;
            }
            QToolButton:hover {
                background-color: #555;
            }
        """)
        self.theme_btn.clicked.connect(self.toggle_theme)
        header_layout.addStretch()
        header_layout.addWidget(self.theme_btn)
        
        left_layout.addLayout(header_layout)
        
        # Symbol input with management
        symbol_group = QGroupBox(self.translator.tr('symbol_management'))
        symbol_layout = QVBoxLayout()
        
        # Symbol input
        symbol_input_layout = QHBoxLayout()
        symbol_input_layout.addWidget(QLabel(self.translator.tr('symbols_label')))
        self.symbol_input = QLineEdit("XAUUSD,EURUSD,GBPJPY,BTCUSD")
        symbol_input_layout.addWidget(self.symbol_input)
        symbol_layout.addLayout(symbol_input_layout)
        
        # Symbol management buttons
        symbol_management_layout = QHBoxLayout()
        
        # Presets combo box
        self.preset_combo = QComboBox()
        self.preset_combo.setPlaceholderText(self.translator.tr('load_preset'))
        self.load_presets()
        self.preset_combo.currentIndexChanged.connect(self.on_preset_selected)
        symbol_management_layout.addWidget(self.preset_combo)
        
        # Save preset button
        self.save_preset_btn = QPushButton(self.translator.tr('save_preset'))
        self.save_preset_btn.clicked.connect(self.save_symbol_preset)
        symbol_management_layout.addWidget(self.save_preset_btn)
        
        # Watchlist button
        self.watchlist_btn = QPushButton(self.translator.tr('load_watchlist'))
        self.watchlist_btn.clicked.connect(self.load_watchlist_symbols)
        symbol_management_layout.addWidget(self.watchlist_btn)
        
        symbol_layout.addLayout(symbol_management_layout)
        symbol_group.setLayout(symbol_layout)
        left_layout.addWidget(symbol_group)

        # Timeframe and export format
        settings_group = QGroupBox(self.translator.tr('settings'))
        settings_layout = QHBoxLayout()
        
        # Timeframe - multi-select list
        tf_frame = QFrame()
        tf_layout = QVBoxLayout()
        tf_layout.addWidget(QLabel(self.translator.tr('timeframes')))
        self.tf_list = QListWidget()
        self.tf_list.setSelectionMode(QListWidget.MultiSelection)
        for tf in ["M1", "M5", "M15", "M30", "H1", "H4", "D1", "W1", "MN1"]:
            item = QListWidgetItem(tf)
            self.tf_list.addItem(item)
        # Select H1 by default
        for i in range(self.tf_list.count()):
            if self.tf_list.item(i).text() == "H1":
                self.tf_list.item(i).setSelected(True)
                break
        self.tf_list.setMaximumHeight(150)
        tf_layout.addWidget(self.tf_list)
        tf_frame.setLayout(tf_layout)
        settings_layout.addWidget(tf_frame)
        
        # Export format
        format_frame = QFrame()
        format_layout = QVBoxLayout()
        format_layout.addWidget(QLabel(self.translator.tr('export_format')))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["xlsx", "csv"])
        format_layout.addWidget(self.format_combo)
        format_frame.setLayout(format_layout)
        settings_layout.addWidget(format_frame)
        
        settings_group.setLayout(settings_layout)
        left_layout.addWidget(settings_group)

        # Date range
        date_group = QGroupBox(self.translator.tr('date_range'))
        date_layout = QVBoxLayout()
        
        # Date selection method
        radio_group = QHBoxLayout()
        self.range_radio = QRadioButton(self.translator.tr('specific_range'))
        self.days_radio = QRadioButton(self.translator.tr('days_back'))
        self.days_radio.setChecked(True)
        self.range_radio.toggled.connect(self.toggle_dates)
        radio_group.addWidget(self.range_radio)
        radio_group.addWidget(self.days_radio)
        date_layout.addLayout(radio_group)
        
        # Date selectors
        self.date_widget = QWidget()
        date_selector_layout = QHBoxLayout()
        date_selector_layout.addWidget(QLabel(self.translator.tr('start_date')))
        self.start_date = QDateEdit(QDate.currentDate().addDays(-30))
        self.start_date.setCalendarPopup(True)
        date_selector_layout.addWidget(self.start_date)
        date_selector_layout.addWidget(QLabel(self.translator.tr('end_date')))
        self.end_date = QDateEdit(QDate.currentDate())
        self.end_date.setCalendarPopup(True)
        date_selector_layout.addWidget(self.end_date)
        self.date_widget.setLayout(date_selector_layout)
        date_layout.addWidget(self.date_widget)
        
        # Days back selector
        self.days_widget = QWidget()
        days_layout = QHBoxLayout()
        days_layout.addWidget(QLabel(self.translator.tr('days_label')))
        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 3650)
        self.days_spin.setValue(30)
        days_layout.addWidget(self.days_spin)
        self.days_widget.setLayout(days_layout)
        date_layout.addWidget(self.days_widget)
        self.toggle_dates()
        
        date_group.setLayout(date_layout)
        left_layout.addWidget(date_group)

        # Output file
        output_group = QGroupBox(self.translator.tr('output'))
        output_layout = QVBoxLayout()
        
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel(self.translator.tr('output_path')))
        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText(self.translator.tr('output_path'))
        file_layout.addWidget(self.file_input)
        self.browse_btn = QPushButton(self.translator.tr('browse_btn'))  # Now stored as instance variable
        self.browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(self.browse_btn)
        output_layout.addLayout(file_layout)
        
        # Column selector
        output_layout.addWidget(QLabel(self.translator.tr('columns_label')))
        self.column_list = QListWidget()
        for col in ["Date", "Open", "High", "Low", "Close", "Volume", "Spread", "RealVolume"]:
            item = QListWidgetItem(col)
            item.setCheckState(Qt.Checked)
            self.column_list.addItem(item)
        self.column_list.setMaximumHeight(150)
        output_layout.addWidget(self.column_list)
        
        output_group.setLayout(output_layout)
        left_layout.addWidget(output_group)

        # Symbol selector for plotting
        self.symbol_combo = QComboBox()
        self.symbol_combo.setPlaceholderText(self.translator.tr('plot_label'))
        left_layout.addWidget(QLabel(self.translator.tr('plot_label')))
        left_layout.addWidget(self.symbol_combo)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        left_layout.addWidget(self.progress_bar)

        # Buttons
        button_layout = QHBoxLayout()
        
        self.download_btn = QPushButton(self.translator.tr('download_btn'))
        self.download_btn.clicked.connect(self.download_data)
        button_layout.addWidget(self.download_btn)
        
        self.stop_btn = QPushButton(self.translator.tr('stop_btn'))
        self.stop_btn.clicked.connect(self.stop_download)
        self.stop_btn.setEnabled(False)
        button_layout.addWidget(self.stop_btn)
        
        self.plot_btn = QPushButton(self.translator.tr('plot_btn'))
        self.plot_btn.clicked.connect(self.plot_selected_symbol)
        self.plot_btn.setEnabled(False)
        button_layout.addWidget(self.plot_btn)
        
        left_layout.addLayout(button_layout)
        
        # Status messages
        self.status_label = QLabel()
        self.status_label.setWordWrap(True)
        left_layout.addWidget(self.status_label)
        
        left_panel.setLayout(left_layout)
        
        # Right panel (chart)
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(10, 10, 10, 10)
        
        self.chart = CandlestickChart(right_panel, width=8, height=6, dpi=100)
        right_layout.addWidget(self.chart)
        
        right_panel.setLayout(right_layout)
        
        # Add panels to splitter
        main_splitter.addWidget(left_panel)
        main_splitter.addWidget(right_panel)
        main_splitter.setStretchFactor(0, 1)
        main_splitter.setStretchFactor(1, 2)
        
        self.setCentralWidget(main_splitter)
        
        # Status bar
        self.statusBar().showMessage("Ready")

    def apply_dark_theme(self):
        """Apply dark theme to all UI elements"""
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(53, 53, 53))
        palette.setColor(QPalette.WindowText, Qt.white)
        palette.setColor(QPalette.Base, QColor(35, 35, 35))
        palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ToolTipBase, QColor(25, 25, 25))
        palette.setColor(QPalette.ToolTipText, Qt.white)
        palette.setColor(QPalette.Text, Qt.white)
        palette.setColor(QPalette.Button, QColor(53, 53, 53))
        palette.setColor(QPalette.ButtonText, Qt.white)
        palette.setColor(QPalette.BrightText, Qt.red)
        palette.setColor(QPalette.Highlight, QColor(142, 45, 197).lighter())
        palette.setColor(QPalette.HighlightedText, Qt.black)
        
        # Apply to buttons
        button_style = """
            QPushButton {
                background-color: #8e2dc5;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #9b3fd4;
            }
            QPushButton:disabled {
                background-color: #555;
                color: #999;
            }
        """
        self.download_btn.setStyleSheet(button_style)
        self.stop_btn.setStyleSheet(button_style)
        self.plot_btn.setStyleSheet(button_style)
        self.save_preset_btn.setStyleSheet(button_style)
        self.watchlist_btn.setStyleSheet(button_style)
        self.lang_btn.setStyleSheet(button_style)
        self.browse_btn.setStyleSheet(button_style)
        
        # Apply to inputs
        input_style = """
            QComboBox, QLineEdit, QSpinBox, QDateEdit {
                background-color: #353535;
                color: white;
                border: 1px solid #666;
                padding: 5px;
                border-radius: 4px;
            }
            QComboBox QAbstractItemView {
                background-color: #353535;
                color: white;
                selection-background-color: #8e2dc5;
            }
        """
        self.symbol_input.setStyleSheet(input_style)
        self.format_combo.setStyleSheet(input_style)
        self.symbol_combo.setStyleSheet(input_style)
        self.days_spin.setStyleSheet(input_style)
        self.start_date.setStyleSheet(input_style)
        self.end_date.setStyleSheet(input_style)
        self.file_input.setStyleSheet(input_style)
        self.preset_combo.setStyleSheet(input_style)
        
        # Apply to list widgets
        list_style = """
            QListWidget {
                background-color: #353535;
                color: white;
                border: 1px solid #666;
                border-radius: 4px;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #8e2dc5;
                color: white;
            }
            QListWidget::item:hover {
                background-color: #666;
            }
        """
        self.tf_list.setStyleSheet(list_style)
        self.column_list.setStyleSheet(list_style)
        
        # Apply to group boxes
        group_style = """
            QGroupBox {
                border: 1px solid #666;
                border-radius: 5px;
                margin-top: 10px;
                color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """
        for widget in self.findChildren(QGroupBox):
            widget.setStyleSheet(group_style)
        
        # Apply to progress bar
        progress_style = """
            QProgressBar {
                border: 1px solid #666;
                border-radius: 4px;
                text-align: center;
                height: 20px;
                color: white;
            }
            QProgressBar::chunk {
                background-color: #8e2dc5;
                width: 10px;
            }
        """
        self.progress_bar.setStyleSheet(progress_style)
        
        # Status label
        self.status_label.setStyleSheet("color: white;")
        
        QApplication.setPalette(palette)
        self.chart.dark_mode = self.dark_mode
        if hasattr(self, 'current_chart_symbol') and self.current_chart_symbol:
            self.plot_data(self.chart_data[self.current_chart_symbol], self.current_chart_symbol)

    def toggle_language(self):
        """Toggle between English and Farsi"""
        try:
            new_lang = 'fa' if self.translator.current_lang == 'en' else 'en'
            self.translator.set_language(new_lang)
            self.retranslate_ui()
        except Exception as e:
            logging.error(f"Error toggling language: {str(e)}")
            QMessageBox.warning(self, "Error", f"Failed to toggle language: {str(e)}")
        
    def retranslate_ui(self):
        """Update all UI elements with current language"""
        self.setWindowTitle(self.translator.tr('app_title'))
        
        # Update group boxes
        for widget in self.findChildren(QGroupBox):
            if widget.title() in self.translator.translations['en'].values():
                for key, value in self.translator.translations['en'].items():
                    if value == widget.title():
                        widget.setTitle(self.translator.tr(key))
                        break
        
        # Update buttons
        self.download_btn.setText(self.translator.tr('download_btn'))
        self.stop_btn.setText(self.translator.tr('stop_btn'))
        self.plot_btn.setText(self.translator.tr('plot_btn'))
        self.save_preset_btn.setText(self.translator.tr('save_preset'))
        self.watchlist_btn.setText(self.translator.tr('load_watchlist'))
        self.browse_btn.setText(self.translator.tr('browse_btn'))
        
        # Update labels
        for widget in self.findChildren(QLabel):
            if widget.text() in self.translator.translations['en'].values():
                for key, value in self.translator.translations['en'].items():
                    if value == widget.text():
                        widget.setText(self.translator.tr(key))
                        break
        
        # Update other elements
        self.preset_combo.setPlaceholderText(self.translator.tr('load_preset'))
        self.symbol_combo.setPlaceholderText(self.translator.tr('plot_label'))
        
        # RTL support for Farsi
        if self.translator.current_lang == 'fa':
            self.setLayoutDirection(Qt.RightToLeft)
            for widget in self.findChildren(QWidget):
                widget.setLayoutDirection(Qt.RightToLeft)
        else:
            self.setLayoutDirection(Qt.LeftToRight)
            for widget in self.findChildren(QWidget):
                widget.setLayoutDirection(Qt.LeftToRight)

    # Symbol Management Methods
    def load_presets(self):
        """Load saved presets from config file"""
        self.preset_combo.clear()
        self.preset_combo.addItem(f"-- {self.translator.tr('load_preset')} --", None)
        try:
            if os.path.exists('symbol_presets.json'):
                with open('symbol_presets.json', 'r') as f:
                    presets = json.load(f)
                    for preset_name in presets.keys():
                        self.preset_combo.addItem(preset_name, presets[preset_name])
        except Exception as e:
            logging.error(f"Error loading presets: {e}")
            self.statusBar().showMessage("Error loading presets", 3000)

    def save_symbol_preset(self):
        """Save current symbols as a named preset"""
        current_symbols = self.symbol_input.text().strip()
        if not current_symbols:
            QMessageBox.warning(self, "Warning", self.translator.tr('no_symbols'))
            return
            
        preset_name, ok = QInputDialog.getText(
            self, self.translator.tr('save_preset'), self.translator.tr('preset_name')
        )
        if ok and preset_name:
            try:
                # Load existing presets
                presets = {}
                if os.path.exists('symbol_presets.json'):
                    with open('symbol_presets.json', 'r') as f:
                        presets = json.load(f)
                
                # Add/update preset
                presets[preset_name] = current_symbols
                
                # Save back to file
                with open('symbol_presets.json', 'w') as f:
                    json.dump(presets, f, indent=4)
                
                self.load_presets()  # Refresh the combo box
                self.statusBar().showMessage(
                    self.translator.tr('preset_saved').format(preset_name), 3000
                )
                
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to save preset: {e}")
                logging.error(f"Error saving preset: {e}")

    def on_preset_selected(self, index):
        """Load symbols from selected preset"""
        if index > 0:  # Skip the first "Select Preset" item
            preset_symbols = self.preset_combo.currentData()
            if preset_symbols:
                self.symbol_input.setText(preset_symbols)
                self.statusBar().showMessage(
                    self.translator.tr('preset_loaded').format(self.preset_combo.currentText()), 
                    3000
                )

    def load_watchlist_symbols(self):
        """Load symbols from MT5 market watch"""
        if not mt5.initialize():
            QMessageBox.warning(self, "Error", self.translator.tr('mt5_connect_error'))
            return
        
        try:
            # Get symbols from market watch
            symbols = mt5.symbols_get()
            watchlist_symbols = [s.name for s in symbols if s.visible]
            
            if not watchlist_symbols:
                QMessageBox.information(self, "Info", self.translator.tr('no_watchlist'))
                return
                
            # Show dialog to select symbols
            dialog = QDialog(self)
            dialog.setWindowTitle(self.translator.tr('watchlist_title'))
            dialog.setMinimumWidth(300)
            layout = QVBoxLayout()
            
            list_widget = QListWidget()
            list_widget.setSelectionMode(QListWidget.MultiSelection)
            for symbol in sorted(watchlist_symbols):
                list_widget.addItem(symbol)
            
            # Select all by default
            for i in range(list_widget.count()):
                list_widget.item(i).setSelected(True)
            
            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            
            layout.addWidget(QLabel(self.translator.tr('select_symbols')))
            layout.addWidget(list_widget)
            layout.addWidget(button_box)
            dialog.setLayout(layout)
            
            if dialog.exec_() == QDialog.Accepted:
                selected_items = [item.text() for item in list_widget.selectedItems()]
                if selected_items:
                    self.symbol_input.setText(",".join(sorted(selected_items)))
                    self.statusBar().showMessage(
                        self.translator.tr('watchlist_loaded').format(len(selected_items)), 
                        3000
                    )
                    
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load watchlist: {e}")
            logging.error(f"Error loading watchlist: {e}")
        finally:
            mt5.shutdown()

    def get_current_symbols(self):
        """Get list of currently entered symbols"""
        return [s.strip() for s in self.symbol_input.text().split(',') if s.strip()]

    def validate_symbols(self, symbols):
        """Validate symbols against MT5's available symbols"""
        if not mt5.initialize():
            return False, self.translator.tr('mt5_connect_error')
        
        try:
            available_symbols = [s.name.lower() for s in mt5.symbols_get()]
            invalid_symbols = [
                s for s in symbols 
                if s.lower() not in available_symbols
            ]
            
            if invalid_symbols:
                return False, f"Invalid symbols: {', '.join(invalid_symbols)}"
            return True, "All symbols valid"
            
        except Exception as e:
            return False, str(e)
        finally:
            mt5.shutdown()

    def toggle_theme(self):
        """Toggle between dark and light themes (dark only as per request)"""
        self.theme_btn.setChecked(True)  # Keep the button in "dark mode" state
        self.apply_dark_theme()
        QMessageBox.information(self, "Theme", self.translator.tr('theme_message'))

    def toggle_dates(self):
        self.date_widget.setVisible(self.range_radio.isChecked())
        self.days_widget.setVisible(self.days_radio.isChecked())

    def browse_file(self):
        fmt = self.format_combo.currentText()
        filter_str = "Excel Files (*.xlsx)" if fmt == "xlsx" else "CSV Files (*.csv)"
        file, _ = QFileDialog.getSaveFileName(self, self.translator.tr('browse_btn'), "", filter_str)
        if file:
            self.file_input.setText(file)

    def plot_selected_symbol(self):
        """Plot the currently selected symbol from combo box"""
        selected_symbol = self.symbol_combo.currentText()
        if not selected_symbol or selected_symbol not in self.chart_data:
            QMessageBox.warning(self, "No Data", "Please select a valid symbol to plot.")
            return
            
        data = self.chart_data[selected_symbol]
        self.current_chart_symbol = selected_symbol
        self.plot_data(data, selected_symbol)

    def plot_data(self, data, symbol=""):
        """Plot data for a specific symbol"""
        try:
            if data is None or data.empty:
                QMessageBox.warning(self, "No Data", self.translator.tr('no_data'))
                return
                
            # We always have OHLC data since it's stored separately
            plot_data = data.copy()
            
            # Convert date if needed
            if not pd.api.types.is_datetime64_any_dtype(plot_data['Date']):
                plot_data['Date'] = pd.to_datetime(plot_data['Date'])
            
            # Plot the data
            self.chart.plot_candles(plot_data, symbol)
            
        except Exception as e:
            QMessageBox.critical(self, "Plot Error", f"Failed to plot data: {str(e)}")
            logging.error(f"Plot error: {str(e)}")

    def download_data(self):
        # Validate symbols
        symbols = self.get_current_symbols()
        if not symbols:
            QMessageBox.warning(self, "Invalid Input", self.translator.tr('no_symbols'))
            return

        # Get selected timeframes
        selected_timeframes = [item.text() for item in self.tf_list.selectedItems()]
        if not selected_timeframes:
            QMessageBox.warning(self, "Invalid Input", self.translator.tr('no_timeframes'))
            return

        # Get date range
        if self.range_radio.isChecked():
            start = self.start_date.date().toPyDate()
            end = self.end_date.date().toPyDate()
            if start > end:
                QMessageBox.warning(self, "Invalid Range", self.translator.tr('invalid_range'))
                return
        else:
            end = datetime.now().date()
            start = end - timedelta(days=self.days_spin.value())
            end = min(end, datetime.now().date())

        # Get other parameters
        export_format = self.format_combo.currentText()
        output_file = self.file_input.text().strip()
        
        # Get selected columns for export (chart always uses OHLC)
        selected_columns = [self.column_list.item(i).text() 
                          for i in range(self.column_list.count())
                          if self.column_list.item(i).checkState() == Qt.Checked]
        if not selected_columns:
            QMessageBox.warning(self, "Column Selection", self.translator.tr('no_columns'))
            return

        # Start download
        self.download_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.plot_btn.setEnabled(False)
        self.statusBar().showMessage(self.translator.tr('download_starting'))
        self.progress_bar.setValue(0)
        self.status_label.setText(self.translator.tr('download_starting'))
        
        self.download_thread = DataDownloadThread(
            symbols, selected_timeframes, 
            datetime.combine(start, datetime.min.time()),
            datetime.combine(end, datetime.min.time()),
            output_file, export_format, selected_columns
        )
        self.download_thread.progress.connect(self.progress_bar.setValue)
        self.download_thread.finished.connect(self.on_download_finished)
        self.download_thread.error.connect(self.download_error)
        self.download_thread.log_message.connect(self.update_status)
        self.download_thread.start()

    def stop_download(self):
        if hasattr(self, 'download_thread') and isinstance(self.download_thread, QThread) and self.download_thread.isRunning():
            self.download_thread.stop()
            self.status_label.setText("Download stopped by user")
            self.download_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.statusBar().showMessage("Download stopped")

    def update_status(self, message, level):
        """Update status label with messages from the download thread"""
        color = {
            "INFO": "white",
            "WARNING": "orange",
            "ERROR": "red"
        }.get(level, "white")
        
        self.status_label.setText(f"<font color='{color}'>{message}</font>")
        self.status_label.repaint()

    def on_download_finished(self, result_data):
        """Handle download completion - result_data contains OHLC data for charting"""
        self.chart_data = result_data
        self.download_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_bar.setValue(100)
        
        # Update symbol combo box
        self.symbol_combo.clear()
        if result_data:
            self.symbol_combo.addItems(sorted(result_data.keys()))
            self.plot_btn.setEnabled(True)
            self.statusBar().showMessage(self.translator.tr('download_complete'))
            self.status_label.setText(self.translator.tr('download_complete'))
            
            # Auto-plot the first symbol/timeframe
            first_key = sorted(result_data.keys())[0]
            self.symbol_combo.setCurrentText(first_key)
            self.current_chart_symbol = first_key
            self.plot_data(result_data[first_key], first_key)
        else:
            self.statusBar().showMessage(self.translator.tr('no_data_received'))
            self.status_label.setText(self.translator.tr('no_data_received'))

    def download_error(self, error_msg):
        self.statusBar().showMessage(self.translator.tr('error_occurred'))
        self.status_label.setText(f"<font color='red'>Error: {error_msg}</font>")
        self.download_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.plot_btn.setEnabled(False)
        self.progress_bar.setValue(0)

    def closeEvent(self, event):
        """Handle window close event to ensure clean shutdown"""
        if hasattr(self, 'download_thread') and isinstance(self.download_thread, QThread) and self.download_thread.isRunning():
            reply = QMessageBox.question(
                self, 'Download in Progress',
                self.translator.tr('confirm_exit'),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.download_thread.stop()
                self.download_thread.wait(1000)  # Wait up to 1 second for thread to finish
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Set application style
    app.setStyleSheet("""
        QToolTip {
            color: #ffffff;
            background-color: #2a2a2a;
            border: 1px solid #444;
            padding: 2px;
        }
    """)
    
    window = MT5DataDownloader()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

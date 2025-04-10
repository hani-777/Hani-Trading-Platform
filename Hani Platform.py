import os
import MetaTrader5 as mt5
import requests
import json
import pandas as pd
from datetime import datetime, timedelta, time as dt_time
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QMessageBox

# MetaTrader 5 initialization
mt5.initialize()

# Define the URL for the webhook
strategy_name = "Hani Trading"
url = f"https://haniwebhook-28e2128c5cfd.herokuapp.com/{strategy_name}"
re = requests.get(url, timeout=10)
# Fixed list of symbols
fixed_symbols = ['XAUUSD.r', 'USDJPY.r', 'USDCAD.r', 'USDCHF.r', 'EURUSD.r', 'AUDUSD.r', 'EURCAD.r', 'EURCHF.r', 'EURGBP.r', 'AUDCAD.r', 'EURJPY.r', 'GBPJPY.r', 'NZDUSD.r', 'XAGUSD.r']
one_time = False
# Global variables for pivot values
last_pivot_high_value = None
last_pivot_low_value = None

# Trade management mode: 1 = Single direction, 2 = Hedging, 3 = All signals, 4 = Smart Hedging
trade_management_mode = 1  # Default value, will be changed by ComboBox

# Define the path for the balance file
balance_file_path = f"daily_balance {mt5.account_info().login}.txt"

def create_balance_file():
    if not os.path.exists(balance_file_path):
        with open(balance_file_path, "w") as file:
            file.write("0\n")  # Initial balance if file doesn't exist

def read_balance_from_file():
    if os.path.exists(balance_file_path):
        with open(balance_file_path, "r") as file:
            return float(file.read().strip())
    else:
        return 0

def write_balance_to_file(balance):
    with open(balance_file_path, "w") as file:
        file.write(f"{balance}\n")

class TradingDashboard(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # Initial values for variables
        self.commission = 10  # commission per lot
        self.tp1 = 50  # TP level 1 percentage
        self.tp2 = 50  # TP level 2 percentage
        self.R1 = 12  # Required profit level 1 as a percentage of balance
        self.R2 = 13  # Required profit level 2 as a percentage of balance
        self.R3 = 15 # Additional required profit level for total profit check as a percentage of balance
        self.daily_profit_target = 10 # Define your daily profit target here as a percentage of balance
        self.trading_stopped = False
        self.first_run = True

        self.create_balance_file()
        self.previous_day_balance = self.read_balance_from_file()

        self.total_profits = {}
        self.high_impact_news = []
        self.auto_trading = True
        self.tp_status = {}
        self.news_management_active = True  # Initialize this before calling get_forex_news
        self.usetotal = True
        self.last_excel_mod_time = None
        self.one_time = False

        self.quiet_hours_start = dt_time(23, 57)  # Start of quiet hours
        self.quiet_hours_end = dt_time(2, 5)  # End of quiet hours and daily update time

        self.symbol_settings = {}

        # Initialize log_display first
        self.log_display = QtWidgets.QPlainTextEdit(self)
        self.log_display.setReadOnly(True)

        # Now, initialize other components and variables
        self.create_balance_file()
        self.previous_day_balance = self.read_balance_from_file()

        # Check current balance and update the file if necessary
        if self.first_run:
            # Load forex news before starting the program
            self.get_forex_news()
            positions = mt5.positions_get()
            current_balance = mt5.account_info().balance
            if len(positions) == 0:
                if current_balance > self.previous_day_balance:
                    self.previous_day_balance = current_balance
                    self.write_balance_to_file(self.previous_day_balance)
            else:
                self.add_log("Cannot update balance, there are open positions.")
            self.first_run = False  # Set the flag to False after the first run

        self.init_ui()
        self.start_daily_timer()

    def apply_tp1_manual(self, ticket):
        """Manually apply TP1 and set the tp1_applied flag."""
        position = next((p for p in mt5.positions_get() if p.ticket == ticket), None)
        if position:
            symbol = position.symbol
            self.partial_close_trade(ticket, self.symbol_settings[symbol]['tp1'])
            
            if ticket not in self.tp_status:
                self.tp_status[ticket] = {'tp1_applied': True, 'tp2_applied': False}
            else:
                self.tp_status[ticket]['tp1_applied'] = True
            
            self.add_log(f"Manual TP1 applied for ticket {ticket}")
        else:
            self.add_log(f"No position found for ticket {ticket} to apply TP1 manually.")

    def apply_tp2_manual(self, ticket):
        """Manually apply TP2, set the tp2_applied flag, and execute break-even."""
        position = next((p for p in mt5.positions_get() if p.ticket == ticket), None)
        if position:
            symbol = position.symbol
            self.partial_close_trade(ticket, self.symbol_settings[symbol]['tp2'])
            self.break_even(ticket)
            
            if ticket not in self.tp_status:
                self.tp_status[ticket] = {'tp1_applied': False, 'tp2_applied': True}
            else:
                self.tp_status[ticket]['tp2_applied'] = True
            
            self.add_log(f"Manual TP2 applied for ticket {ticket} and break-even executed.")
        else:
            self.add_log(f"No position found for ticket {ticket} to apply TP2 manually.")

    def create_balance_file(self):
        create_balance_file()

    def read_balance_from_file(self):
        return read_balance_from_file()

    def write_balance_to_file(self, balance):
        write_balance_to_file(balance)

    def init_ui(self):
        self.setWindowTitle(f"Hani Trading Dashboard,     Platform: {strategy_name},      Account Number: {mt5.account_info().login}")
        self.setGeometry(100, 100, 1920, 1080)

        layout = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableWidget(self)
        self.table.setColumnCount(18)
        self.table.setHorizontalHeaderLabels([
            'Keys', 'Ticket', 'Symbol', 'Type', 'Volume', 'Open Price', 'Current Price',
            'Stop Loss', 'Take Profit', 'Commission', 'Total Profit', 'Profit', 'Real Profit',
            'Real Profit(%)', 'TP1 Level', 'TP2 Level', 'Full Close', 'Real SL'
        ])
        layout.addWidget(self.table)

        # Set column widths
        self.table.setColumnWidth(0, 280)  # Keys (adjust as needed for action buttons)
        self.table.setColumnWidth(1, 120)  # Ticket
        self.table.setColumnWidth(2, 120)  # Symbol
        self.table.setColumnWidth(3, 80)   # Type
        self.table.setColumnWidth(4, 60)   # Volume
        self.table.setColumnWidth(5, 110)  # Open Price
        self.table.setColumnWidth(6, 110)  # Current Price
        self.table.setColumnWidth(7, 110)  # Stop Loss
        self.table.setColumnWidth(8, 110)  # Take Profit
        self.table.setColumnWidth(9, 80)   # Commission
        self.table.setColumnWidth(10, 90)  # Total Profit
        self.table.setColumnWidth(11, 90)  # Profit
        self.table.setColumnWidth(12, 90)  # Real Profit
        self.table.setColumnWidth(13, 95)  # Real Profit(%)
        self.table.setColumnWidth(14, 80)  # TP1 Level
        self.table.setColumnWidth(15, 80)  # TP2 Level
        self.table.setColumnWidth(16, 80)  # Full Close
        self.table.setColumnWidth(17, 80)  # Real SL

        controls_layout = QtWidgets.QHBoxLayout()
        controls_layout.setSpacing(5)  # Add spacing between elements
        controls_layout.setContentsMargins(5, 5, 5, 5)  # Set margins

        self.required_profit_label = QtWidgets.QLabel("Required Profit to Close all: 0.00", self)

        self.buy_price_label = QtWidgets.QLabel("|                                             Buy Price: 0.00", self)

        self.sell_price_label = QtWidgets.QLabel("|                                            Sell Price: 0.00", self)
        
        self.total_real_profit_label = QtWidgets.QLabel("Total Real Profit: 0.00", self)
        self.total_real_profit_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignRight)

        price_layout = QtWidgets.QHBoxLayout()
        price_layout.addWidget(self.required_profit_label)
        price_layout.addWidget(self.buy_price_label)
        price_layout.addWidget(self.sell_price_label)
        price_layout.addWidget(self.total_real_profit_label)

        self.price_group_box = QtWidgets.QGroupBox("Important Data")
        self.price_group_box.setLayout(price_layout)

        layout.addWidget(self.price_group_box)


        self.update_daily_profit_button = QtWidgets.QPushButton("Change D P T", self)
        self.update_daily_profit_button.setStyleSheet("background-color: grey; color: white;")
        self.update_daily_profit_button.clicked.connect(self.update_daily_profit_target)
        self.update_daily_profit_button.setMaximumWidth(100)
        controls_layout.addWidget(self.update_daily_profit_button)

        self.add_symbol_button = QtWidgets.QPushButton("Add/Update", self)
        self.add_symbol_button.setStyleSheet("background-color: grey; color: white;")
        self.add_symbol_button.clicked.connect(self.add_symbol)
        self.add_symbol_button.setMaximumWidth(100)
        controls_layout.addWidget(self.add_symbol_button)

        self.symbol_combobox = QtWidgets.QComboBox(self)
        self.symbol_combobox.addItems(fixed_symbols)
        font = self.symbol_combobox.font()
        font.setPointSize(8)
        self.symbol_combobox.setFont(font)
        self.symbol_combobox.setMaximumWidth(140)
        self.symbol_combobox.view().setSpacing(3)

        controls_layout.addWidget(self.symbol_combobox)

        # Add the control buttons for new features
        self.close_all_button = QtWidgets.QPushButton("Close All", self)
        self.close_all_button.setStyleSheet("background-color: orange; color: white;")
        self.close_all_button.clicked.connect(self.close_all_positions)
        self.close_all_button.setMaximumWidth(140)
        controls_layout.addWidget(self.close_all_button)

        # Add the "Clear All Orders" button
        self.clear_all_orders_button = QtWidgets.QPushButton("Clear All Orders", self)
        self.clear_all_orders_button.setStyleSheet("background-color: grey; color: white;")
        self.clear_all_orders_button.clicked.connect(self.clear_all_orders)
        self.clear_all_orders_button.setMaximumWidth(150)
        controls_layout.addWidget(self.clear_all_orders_button)

        self.close_all_profit_button = QtWidgets.QPushButton("Close All on Profit", self)
        self.close_all_profit_button.setStyleSheet("background-color: green; color: white;")
        self.close_all_profit_button.clicked.connect(self.close_all_in_profit)
        self.close_all_profit_button.setMaximumWidth(160)
        controls_layout.addWidget(self.close_all_profit_button)

        self.close_all_loss_button = QtWidgets.QPushButton("Close All in Loss", self)
        self.close_all_loss_button.setStyleSheet("background-color: red; color: white;")
        self.close_all_loss_button.clicked.connect(self.close_all_in_loss)
        self.close_all_loss_button.setMaximumWidth(160)
        controls_layout.addWidget(self.close_all_loss_button)

        self.auto_trading_button = QtWidgets.QPushButton("Auto Trading: On", self)
        self.auto_trading_button.setStyleSheet("background-color: green; color: white;")
        self.auto_trading_button.clicked.connect(self.toggle_auto_trading)
        self.auto_trading_button.setMaximumWidth(120)
        controls_layout.addWidget(self.auto_trading_button)

        self.news_management_button = QtWidgets.QPushButton("News Management: on", self)
        self.news_management_button.setStyleSheet("background-color: green; color: white;")
        self.news_management_button.clicked.connect(self.toggle_news_management)
        self.news_management_button.setMaximumWidth(165)
        controls_layout.addWidget(self.news_management_button)

        self.usetotal_button = QtWidgets.QPushButton("Use Total Profit: On", self)
        self.usetotal_button.setStyleSheet("background-color: green; color: white;")
        self.usetotal_button.clicked.connect(self.toggle_usetotal)
        self.usetotal_button.setMaximumWidth(150)
        controls_layout.addWidget(self.usetotal_button)

        self.reset_total_profit_button = QtWidgets.QPushButton("Reset Total Profit", self)
        self.reset_total_profit_button.setStyleSheet("background-color: orange; color: black;")
        self.reset_total_profit_button.clicked.connect(self.reset_total_profit)
        self.reset_total_profit_button.setMaximumWidth(150)
        controls_layout.addWidget(self.reset_total_profit_button)

        self.trade_mode_combobox = QtWidgets.QComboBox(self)
        self.trade_mode_combobox.addItems(["Single Direction", "Hedging", "All Signals", "Smart Hedging"])
        self.trade_mode_combobox.setCurrentIndex(trade_management_mode - 1)
        self.trade_mode_combobox.currentIndexChanged.connect(self.update_trade_management_mode)
        self.trade_mode_combobox.setMaximumWidth(160)
        controls_layout.addWidget(self.trade_mode_combobox)

        layout.addLayout(controls_layout)

        self.variables_layout = QtWidgets.QGridLayout()
        self.variables_layout.setSpacing(5)
        self.variables_layout.setContentsMargins(5, 5, 5, 5)

        self.variables_layout.addWidget(QtWidgets.QLabel("Daily Profit Target (%):"), 0, 0)
        self.daily_profit_target_entry = QtWidgets.QDoubleSpinBox(self)
        self.daily_profit_target_entry.setValue(self.daily_profit_target)
        self.daily_profit_target_entry.setMaximum(100000.00)
        self.daily_profit_target_entry.setFixedWidth(100)
        self.daily_profit_target_entry.setMaximumWidth(200)
        self.daily_profit_target_entry.setToolTip("Daily profit target to stop trading for All symbol")
        self.variables_layout.addWidget(self.daily_profit_target_entry, 0, 1)

        spacer_item = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum)
        self.variables_layout.addItem(spacer_item, 0, 2)  # Add spacer between widgets

        self.variables_layout.addWidget(QtWidgets.QLabel("TP1 (%):"), 0, 3)
        self.tp1_entry = QtWidgets.QDoubleSpinBox(self)
        self.tp1_entry.setValue(self.tp1)
        self.tp1_entry.setMinimumWidth(100)
        self.tp1_entry.setMaximumWidth(200)
        self.tp1_entry.setToolTip("Take Profit level 1 percentage for All symbol")
        self.variables_layout.addWidget(self.tp1_entry, 0, 4)

        spacer_item = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum)
        self.variables_layout.addItem(spacer_item, 0, 5)  # Add spacer between widgets

        self.variables_layout.addWidget(QtWidgets.QLabel("TP2 (%):"), 0, 6)
        self.tp2_entry = QtWidgets.QDoubleSpinBox(self)
        self.tp2_entry.setValue(self.tp2)
        self.tp2_entry.setMinimumWidth(100)
        self.tp2_entry.setMaximumWidth(200)
        self.tp2_entry.setToolTip("Take Profit level 2 percentage for All symbol")
        self.variables_layout.addWidget(self.tp2_entry, 0, 7)

        spacer_item_1 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum)
        self.variables_layout.addItem(spacer_item_1, 0, 8)  # Add spacer between widgets

        self.variables_layout.addWidget(QtWidgets.QLabel("Commission:"), 0, 9)
        self.commission_entry = QtWidgets.QDoubleSpinBox(self)
        self.commission_entry.setValue(self.commission)
        self.commission_entry.setMaximumWidth(140)
        self.commission_entry.setFixedWidth(100)
        self.commission_entry.setToolTip("Commission per lot for each symbol")
        self.variables_layout.addWidget(self.commission_entry, 0, 10)

        spacer_item_2 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.variables_layout.addItem(spacer_item_2, 0, 11)  # Add spacer between widgets

        self.variables_layout.addWidget(QtWidgets.QLabel("R1 (%):"), 0, 12)
        self.R1_entry = QtWidgets.QDoubleSpinBox(self)
        self.R1_entry.setValue(self.R1)
        self.R1_entry.setMaximum(100.00)
        self.R1_entry.setMinimumWidth(100)
        self.R1_entry.setMaximumWidth(200)
        self.R1_entry.setToolTip("Required profit level 1 as a percentage of balance for each symbol")
        self.variables_layout.addWidget(self.R1_entry, 0, 13)

        spacer_item_3 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.variables_layout.addItem(spacer_item_3, 0, 14)

        self.variables_layout.addWidget(QtWidgets.QLabel("R2 (%):"), 0, 15)
        self.R2_entry = QtWidgets.QDoubleSpinBox(self)
        self.R2_entry.setValue(self.R2)
        self.R2_entry.setMaximum(100.00)
        self.R2_entry.setMinimumWidth(100)
        self.R2_entry.setMaximumWidth(200)
        self.R2_entry.setToolTip("Required profit level 2 as a percentage of balance for each symbol")
        self.variables_layout.addWidget(self.R2_entry, 0, 16)

        spacer_item_4 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.variables_layout.addItem(spacer_item_4, 0, 17)

        self.variables_layout.addWidget(QtWidgets.QLabel("R3 (%):"), 0, 18)
        self.R3_entry = QtWidgets.QDoubleSpinBox(self)
        self.R3_entry.setValue(self.R3)
        self.R3_entry.setMaximum(100.00)
        self.R3_entry.setMinimumWidth(100)
        self.R3_entry.setMaximumWidth(200)
        self.R3_entry.setToolTip("Additional required profit level for total profit check as a percentage of balance for each symbol")
        self.variables_layout.addWidget(self.R3_entry, 0, 19)

        layout.addLayout(self.variables_layout)

        self.news_info_label = QtWidgets.QLabel("News Info: No upcoming news", self)
        layout.addWidget(self.news_info_label)

        self.time_to_news_label = QtWidgets.QLabel("Time to News: N/A", self)
        layout.addWidget(self.time_to_news_label)

        self.time_after_news_label = QtWidgets.QLabel("Time After News: N/A", self)
        layout.addWidget(self.time_after_news_label)

        self.quiet_hours_label = QtWidgets.QLabel("Quiet Hours: Trading hours active", self)
        layout.addWidget(self.quiet_hours_label)

        # Create a layout for manual trading buttons
        self.manual_trade_layout = QtWidgets.QVBoxLayout()
        self.manual_trade_buttons = {}
        layout.addLayout(self.manual_trade_layout)

        # Log Display Widget
        self.log_display = QtWidgets.QPlainTextEdit(self)
        self.log_display.setReadOnly(True)
        self.log_display.setFixedHeight(120)
        layout.addWidget(self.log_display)

        self.setLayout(layout)

    def clear_all_orders(self):
        """Clear all pending orders."""
        try:
            orders = mt5.orders_get()
            if orders:
                for order in orders:
                    result = mt5.order_send(
                        action=mt5.TRADE_ACTION_REMOVE,
                        order=order.ticket,
                        symbol=order.symbol,
                        type=order.type
                    )
                    if result.retcode == mt5.TRADE_RETCODE_DONE:
                        self.add_log(f"Successfully removed order {order.symbol}, ticket {order.ticket}")
                    else:
                        self.add_log(f"Failed to remove order {order.symbol}, ticket {order.ticket}: {result.comment}")
            else:
                self.add_log("No pending orders found.")
        except Exception as e:
            self.add_log(f"Error in removing orders: {e}")

    def close_all_in_profit(self):
        self.add_log(f"Attempting to close all profitable positions...")
        positions = mt5.positions_get()
        if positions:
            for pos in positions:
                if pos.profit > 0:
                    self.close_position(pos)
            self.add_log(f"All profitable positions closed.")
        else:
            self.add_log("No open positions found.")

    def close_all_in_loss(self):
        self.add_log(f"Attempting to close all losing positions...")
        positions = mt5.positions_get()
        if positions:
            for pos in positions:
                if pos.profit < 0:
                    self.close_position(pos)
            self.add_log(f"All losing positions closed.")
        else:
            self.add_log("No open positions found.")

    def update_trade_management_mode(self):
        global trade_management_mode
        trade_management_mode = self.trade_mode_combobox.currentIndex() + 1
        self.add_log(f"Trade management mode updated to {trade_management_mode}")

    def toggle_auto_trading(self):
        self.auto_trading = not self.auto_trading
        self.auto_trading_button.setText("Auto Trading: On" if self.auto_trading else "Auto Trading: Off")
        self.auto_trading_button.setStyleSheet("background-color: green; color: white;" if self.auto_trading else "background-color: red; color: white;")
        self.add_log(f"Auto Trading {'enabled' if self.auto_trading else 'disabled'}")

    def toggle_news_management(self):
        self.news_management_active = not self.news_management_active
        self.news_management_button.setText("News Management: On" if self.news_management_active else "News Management: Off")
        self.news_management_button.setStyleSheet("background-color: green; color: white;" if self.news_management_active else "background-color: red; color: white;")
        self.add_log(f"News management is now {'active' if self.news_management_active else 'inactive'}")

    def toggle_usetotal(self):
        self.usetotal = not self.usetotal
        self.usetotal_button.setText("Use Total Profit: On" if self.usetotal else "Use Total Profit: Off")
        self.usetotal_button.setStyleSheet("background-color: green; color: white;" if self.usetotal else "background-color: red; color: white;")
        self.add_log(f"Total Profit usage is now {'enabled' if self.usetotal else 'disabled'}")

    def reset_total_profit(self):
        self.total_profits = {symbol: 0 for symbol in self.total_profits}
        self.add_log("Total Profit has been reset to zero for all symbols")
        self.update_gui()

    def update_variables(self):
        current_symbol = self.symbol_combobox.currentText()
        self.symbol_settings[current_symbol] = {
            'commission': self.commission_entry.value(),
            'tp1': self.tp1_entry.value(),
            'tp2': self.tp2_entry.value(),
            'R1': self.R1_entry.value(),
            'R2': self.R2_entry.value(),
            'R3': self.R3_entry.value()
        }
        self.add_log(f"Updated variables for {current_symbol}: {self.symbol_settings[current_symbol]}")

    def update_daily_profit_target(self):
        self.daily_profit_target = self.daily_profit_target_entry.value()
        self.add_log(f"Updated daily profit target: {self.daily_profit_target}")

    def manual_buy(self):
        symbol = self.symbol_combobox.currentText()
        volume = float(self.volume_entry.text())
        self.open_position(symbol, volume, "Buy")

    def manual_sell(self):
        symbol = self.symbol_combobox.currentText()
        volume = float(self.volume_entry.text())
        self.open_position(symbol, volume, "Sell")

    def open_position(self, symbol, lot, direction, retries=5, delay=60000):
        attempt = 0

        def attempt_open():
            nonlocal attempt
            try:
                if direction == "Buy":
                    result = mt5.order_send(
                        action=mt5.TRADE_ACTION_DEAL,
                        symbol=symbol,
                        volume=lot,
                        type=mt5.ORDER_TYPE_BUY,
                        price=mt5.symbol_info_tick(symbol).ask,
                        deviation=20,
                        magic=234000,
                        comment=strategy_name,
                        type_time=mt5.ORDER_TIME_GTC,
                        type_filling=mt5.ORDER_FILLING_IOC
                    )
                    if result.retcode == mt5.TRADE_RETCODE_DONE:
                        self.add_log(f"Buy trade executed successfully for {symbol} with lot size {lot}")
                    else:
                        self.add_log(f"Failed to execute buy trade for {symbol}: {result.comment}")
                        retry()

                elif direction == "Sell":
                    result = mt5.order_send(
                        action=mt5.TRADE_ACTION_DEAL,
                        symbol=symbol,
                        volume=lot,
                        type=mt5.ORDER_TYPE_SELL,
                        price=mt5.symbol_info_tick(symbol).bid,
                        deviation=20,
                        magic=234000,
                        comment=strategy_name,
                        type_time=mt5.ORDER_TIME_GTC,
                        type_filling=mt5.ORDER_FILLING_IOC
                    )
                    if result.retcode == mt5.TRADE_RETCODE_DONE:
                        self.add_log(f"Sell trade executed successfully for {symbol} with lot size {lot}")
                    else:
                        self.add_log(f"Failed to execute sell trade for {symbol}: {result.comment}")
                        retry()

            except Exception as e:
                self.add_log(f"Error in opening position: {e}")
                retry()

        def retry():
            nonlocal attempt
            attempt += 1
            if attempt < retries:
                self.add_log(f"Retrying to open position for {symbol} (Attempt {attempt + 1}/{retries})...")
                QtCore.QTimer.singleShot(delay, attempt_open)

        attempt_open()

    def close_position(self, position, retries=5, delay=60000):
        attempt = 0

        def attempt_close():
            nonlocal attempt
            try:
                result = mt5.order_send(
                    action=mt5.TRADE_ACTION_DEAL,
                    symbol=position.symbol,
                    volume=position.volume,
                    type=mt5.ORDER_TYPE_BUY if position.type == mt5.ORDER_TYPE_SELL else mt5.ORDER_TYPE_SELL,
                    position=position.ticket,
                    price=mt5.symbol_info_tick(position.symbol).bid if position.type == mt5.ORDER_TYPE_BUY else mt5.symbol_info_tick(position.symbol).ask,
                    deviation=20,
                    magic=234000,
                    comment="Closing position",
                    type_time=mt5.ORDER_TIME_GTC,
                    type_filling=mt5.ORDER_FILLING_IOC
                )
                if result.retcode == mt5.TRADE_RETCODE_DONE:
                    self.add_log(f"Closed position for {position.symbol} with profit {position.profit}")
                else:
                    self.add_log(f"Failed to close position for {position.symbol}: {result.comment}")
                    retry()

            except Exception as e:
                self.add_log(f"Error in closing position: {e}")
                retry()

        def retry():
            nonlocal attempt
            attempt += 1
            if attempt < retries:
                self.add_log(f"Retrying to close position for {position.symbol} (Attempt {attempt + 1}/{retries})...")
                QtCore.QTimer.singleShot(delay, attempt_close)

        attempt_close()

    def apply_tp_logic(self, symbol, lot):
        positions = mt5.positions_get(symbol=symbol)
        if positions:
            for pos in positions:
                ticket = pos.ticket
                profit = pos.profit

                real_profit = self.total_profits.get(symbol, 0) + sum([p.profit for p in positions if p.symbol == symbol])

                # TP status for the current trade
                if ticket not in self.tp_status:
                    self.tp_status[ticket] = {'tp1_applied': False, 'tp2_applied': False}

                tp1_applied = self.tp_status[ticket]['tp1_applied']
                tp2_applied = self.tp_status[ticket]['tp2_applied']

                current_balance = mt5.account_info().balance
                symbol_settings = self.symbol_settings.get(symbol, {
                    'commission': self.commission,
                    'tp1': self.tp1,
                    'tp2': self.tp2,
                    'R1': self.R1,
                    'R2': self.R2,
                    'R3': self.R3,
                    'loss_threshold': 0.0
                })

                # Use stored settings
                tp1_percentage = symbol_settings['tp1']
                tp2_percentage = symbol_settings['tp2']
                R1 = symbol_settings['R1']
                R2 = symbol_settings['R2']
                R3 = symbol_settings['R3']
                commission = symbol_settings['commission']
                loss_threshold = symbol_settings.get('loss_threshold', 0.0)

                # Partial close if profit exceeds R1
                if profit > (R1 * current_balance / 100) + (lot * commission) and not tp1_applied:
                    self.add_log(f"Applying TP1: Closing {tp1_percentage}% of position for ticket {ticket}")
                    self.partial_close_trade(pos.ticket, tp1_percentage)
                    self.tp_status[ticket]['tp1_applied'] = True

                # Partial close if real profit exceeds R2
                if real_profit > (R2 * current_balance / 100) + (lot * commission) and not tp2_applied:
                    self.add_log(f"Applying TP2: Closing {tp2_percentage}% of position for ticket {ticket}")
                    self.partial_close_trade(pos.ticket, tp2_percentage)
                    self.tp_status[ticket]['tp2_applied'] = True
                    self.break_even(ticket)

                # Close entire position if real profit exceeds R3 and total profit is positive
                if real_profit > (R3 * current_balance / 100) and self.total_profits.get(symbol, 0) > 0:
                    self.add_log(f"Closing entire position for {symbol} as real profit exceeds R3")
                    for p in positions:
                        self.close_position(p)
                    self.add_log(f"Updated total profit for {symbol} after closing: {self.total_profits.get(symbol, 0)}")

                # Check for loss threshold
                if loss_threshold > 0:
                    # Compute loss percentage based on account balance
                    position_value = mt5.account_info().balance
                    if pos.profit < 0:
                        loss_percentage = (-pos.profit / position_value) * 100
                        if tp1_applied:
                            loss_threshold = loss_threshold / 2

                        # Check if the loss percentage exceeds the threshold
                        if loss_percentage >= loss_threshold:
                            self.add_log(f"Position {pos.ticket} for {pos.symbol} reached loss threshold of {loss_threshold}%, closing position.")
                            self.close_position(pos)

                self.update_required_profit_label()
        # Remove closed positions from tp_status
        for pos_ticket in list(self.tp_status.keys()):
            if not any(p.ticket == pos_ticket for p in mt5.positions_get()):
                self.add_log(f"Removing closed position ticket {pos_ticket} from tp_status")
                del self.tp_status[pos_ticket]
                positions = mt5.positions_get(symbol=symbol)
                if not positions:
                    self.total_profits[symbol] = 0

    def close_trade(self, ticket):
        position = mt5.positions_get(ticket=ticket)
        if position:
            self.close_position(position[0])

    def partial_close_trade(self, ticket, tp_percentage):
        try:
            # Get the position by ticket
            position = next((pos for pos in mt5.positions_get() if pos.ticket == ticket), None)
            if position:
                # Calculate the volume to close based on tp_percentage
                volume_to_close = round(position.volume * (tp_percentage / 100.0), 2)

                # Close the calculated volume
                self.partial_close_position(position, volume_to_close)
            else:
                self.add_log(f"No position found for ticket {ticket}")

        except Exception as e:
            self.add_log(f"Error in partial closing trade: {e}")

    def partial_close_position(self, position, volume_to_close):
        try:
            # Check if volume to close is greater than or equal to current volume
            if volume_to_close >= position.volume:
                self.add_log(f"Volume to close ({volume_to_close}) is greater than or equal to position volume ({position.volume}). Closing the entire position.")
                volume_to_close = position.volume  # Adjust volume to close to the entire position volume

            # Save current position profit before closing
            current_profit = position.profit

            result = mt5.order_send(
                action=mt5.TRADE_ACTION_DEAL,
                symbol=position.symbol,
                volume=volume_to_close,
                type=mt5.ORDER_TYPE_BUY if position.type == mt5.ORDER_TYPE_SELL else mt5.ORDER_TYPE_SELL,
                position=position.ticket,
                price=mt5.symbol_info_tick(position.symbol).bid if position.type == mt5.ORDER_TYPE_BUY else mt5.symbol_info_tick(position.symbol).ask,
                deviation=20,
                magic=234000,
                comment="Partial Close",
                type_time=mt5.ORDER_TIME_GTC,
                type_filling=mt5.ORDER_FILLING_IOC
            )
            if result.retcode == mt5.TRADE_RETCODE_DONE:
                self.add_log(f"Successfully closed {volume_to_close} volume for {position.symbol}")
                # Update total profit using current profit before closing
                symbol_settings = self.symbol_settings.get(position.symbol, {'commission': self.commission})
                commission = symbol_settings['commission']
                if position.symbol in self.total_profits:
                    self.total_profits[position.symbol] += (current_profit - (volume_to_close * commission)) / position.volume * volume_to_close
                else:
                    self.total_profits[position.symbol] = round((current_profit - (volume_to_close * commission)) / position.volume * volume_to_close, 2)
                
                self.add_log(f"Updated total profit for {position.symbol}: {self.total_profits[position.symbol]}")
            else:
                self.add_log(f"Failed to close position for {position.symbol}: {result.comment}")
        except Exception as e:
            self.add_log(f"Error in partial closing trade: {e}")

    def is_time_between(self, start_time, end_time, current_time=None):
        if current_time is None:
            current_time = datetime.now().time()
        if start_time < end_time:
            return start_time <= current_time <= end_time
        else:
            return start_time <= current_time or current_time <= end_time

    def check_trading_hours(self):
        current_time = datetime.now().time()
        
        if self.is_time_between(self.quiet_hours_start, self.quiet_hours_end):
            self.quiet_hours_label.setText(f"Quiet Hours: Trading off from {self.quiet_hours_start.strftime('%H:%M')} to {self.quiet_hours_end.strftime('%H:%M')}")
            self.manage_trades_during_quiet_hours()
            return False
        else:
            self.quiet_hours_label.setText("Quiet Hours: Trading hours active")
            return True

    def start_daily_timer(self):
        current_time = datetime.now()
        daily_time = datetime.combine(current_time.date(), self.quiet_hours_end)
        if current_time > daily_time:
            daily_time += timedelta(days=1)
        
        delay = (daily_time - current_time).total_seconds()

        QtCore.QTimer.singleShot(int(delay * 1000), self.daily_update)

    def daily_update(self):
        self.add_log("Daily update triggered, checking for Excel modifications...")
        self.get_forex_news()
        self.read_news_from_excel()

        self.resume_trading_next_day()
        self.start_daily_timer()

    def update_total_real_profit(self):
        total_real_profit = 0.0
        positions = mt5.positions_get()
        if positions:
            for pos in positions:
                symbol_settings = self.symbol_settings.get(pos.symbol, {'commission': self.commission})
                commission = symbol_settings['commission']
                real_profit = self.total_profits.get(pos.symbol, 0) + pos.profit - (pos.volume * commission)
                total_real_profit += real_profit
        
        total_real_profit_percentage = round((total_real_profit / mt5.account_info().balance) * 100,2)
        self.total_real_profit_label.setText(f" Total Real Profit{total_real_profit:.2f} USD ({total_real_profit_percentage:.2f}%)")

    def update_gui(self):
        positions = mt5.positions_get()
        self.table.setRowCount(0)
        if positions:
            self.table.setRowCount(len(positions))
            for row, pos in enumerate(positions):
                symbol_info = mt5.symbol_info(pos.symbol)
                if symbol_info is None:
                    continue

                # Get symbol settings
                symbol_settings = self.symbol_settings.get(pos.symbol, {
                    'commission': self.commission,
                    'tp1': self.tp1,
                    'tp2': self.tp2,
                    'R1': self.R1,
                    'R2': self.R2,
                    'R3': self.R3,
                    'loss_threshold': 0.0
                })
                commission = symbol_settings['commission']
                R1 = symbol_settings['R1']
                R2 = symbol_settings['R2']
                lot = pos.volume
                current_balance = mt5.account_info().balance

                # Calculate Real Profit and Real Profit Percentage
                real_profit = self.total_profits.get(pos.symbol, 0) + pos.profit - (pos.volume * commission)
                real_profit_percentage = round((real_profit / mt5.account_info().balance) * 100, 2)

                # Calculate TP1 and TP2 Levels
                TP1_Level = round((R1 * current_balance / 100) + (lot * commission), 2)
                TP2_Level = round((R2 * current_balance / 100) + (lot * commission), 2)

                # Column 0: 'Keys' - action buttons
                actions_widget = QtWidgets.QWidget()
                actions_layout = QtWidgets.QHBoxLayout(actions_widget)
                actions_layout.setContentsMargins(0, 0, 0, 0)
                actions_layout.setSpacing(5)

                close_button = QtWidgets.QPushButton(f"✖")
                close_button.setStyleSheet("background-color: red; color: white;")
                close_button.setFixedSize(25, 20)
                close_button.clicked.connect(lambda _, t=pos.ticket: self.close_trade(t))
                actions_layout.addWidget(close_button)

                # Modify TP1 Button
                tp1_button = QtWidgets.QPushButton(f"TP1")
                tp1_button.setStyleSheet("background-color: green; color: white;")
                tp1_button.setFixedSize(25, 20)
                tp1_button.clicked.connect(lambda _, t=pos.ticket: self.apply_tp1_manual(t))
                actions_layout.addWidget(tp1_button)

                # Modify TP2 Button
                tp2_button = QtWidgets.QPushButton(f"TP2")
                tp2_button.setStyleSheet("background-color: yellow; color: black;")
                tp2_button.setFixedSize(25, 20)
                tp2_button.clicked.connect(lambda _, t=pos.ticket: self.apply_tp2_manual(t))
                actions_layout.addWidget(tp2_button)

                break_even_button = QtWidgets.QPushButton(f"BE")
                break_even_button.setStyleSheet("background-color: blue; color: white;")
                break_even_button.setFixedSize(25, 20)
                break_even_button.clicked.connect(lambda _, t=pos.ticket: self.break_even(t))
                actions_layout.addWidget(break_even_button)

                reset_profit_button = QtWidgets.QPushButton(f"Reset-P")
                reset_profit_button.setStyleSheet("background-color: blue; color: white;")
                reset_profit_button.setFixedSize(55, 20)
                reset_profit_button.clicked.connect(lambda _, s=pos.symbol: self.reset_symbol_profit(s))
                actions_layout.addWidget(reset_profit_button)

                reverse_button = QtWidgets.QPushButton(f"REVERSE")
                reverse_button.setStyleSheet("background-color: purple; color: white;")
                reverse_button.setFixedSize(55, 20)
                reverse_button.clicked.connect(lambda _, t=pos.ticket: self.reverse_trade(t))
                actions_layout.addWidget(reverse_button)

                self.table.setCellWidget(row, 0, actions_widget)

                # Column 1: 'Ticket'
                self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(pos.ticket)))

                # Column 2: 'Symbol'
                self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(pos.symbol))

                # Column 3: 'Type'
                self.table.setItem(row, 3, QtWidgets.QTableWidgetItem("Buy" if pos.type == mt5.ORDER_TYPE_BUY else "Sell"))

                # Column 4: 'Volume'
                self.table.setItem(row, 4, QtWidgets.QTableWidgetItem(str(pos.volume)))

                # Column 5: 'Open Price'
                self.table.setItem(row, 5, QtWidgets.QTableWidgetItem(str(round(pos.price_open, symbol_info.digits))))

                # Column 6: 'Current Price'
                current_price = mt5.symbol_info_tick(pos.symbol).bid if pos.type == mt5.ORDER_TYPE_BUY else mt5.symbol_info_tick(pos.symbol).ask
                self.table.setItem(row, 6, QtWidgets.QTableWidgetItem(str(round(current_price, symbol_info.digits))))

                # Column 7: 'Stop Loss' (Actual SL placed)
                if pos.sl not in (0.0, None):
                    sl_value = str(round(pos.sl, symbol_info.digits))
                else:
                    sl_value = '0'
                self.table.setItem(row, 7, QtWidgets.QTableWidgetItem(sl_value))
                self.table.item(row, 7).setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

                # Column 8: 'Take Profit'
                if pos.tp not in (0.0, None):
                    tp_value = str(round(pos.tp, symbol_info.digits))
                else:
                    tp_value = '0'
                self.table.setItem(row, 8, QtWidgets.QTableWidgetItem(tp_value))
                self.table.item(row, 8).setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

                # Column 9: 'Commission'
                self.table.setItem(row, 9, QtWidgets.QTableWidgetItem(str(round(pos.volume * commission, 2))))

                # Column 10: 'Total Profit'
                self.table.setItem(row, 10, QtWidgets.QTableWidgetItem(str(round(self.total_profits.get(pos.symbol, 0), 2))))

                # Column 11: 'Profit'
                self.table.setItem(row, 11, QtWidgets.QTableWidgetItem(str(round(pos.profit, 2))))

                # Column 12: 'Real Profit'
                self.table.setItem(row, 12, QtWidgets.QTableWidgetItem(str(round(real_profit, 2))))

                # Column 13: 'Real Profit(%)'
                self.table.setItem(row, 13, QtWidgets.QTableWidgetItem(f"{real_profit_percentage:.2f}%"))

                # Column 14: 'TP1 Level'
                self.table.setItem(row, 14, QtWidgets.QTableWidgetItem(str(TP1_Level)))

                # Column 15: 'TP2 Level'
                self.table.setItem(row, 15, QtWidgets.QTableWidgetItem(str(TP2_Level)))

                # Column 16: 'Full Close'
                full_close_value = round((symbol_settings['R3'] * current_balance / 100) + (lot * commission), 2)
                self.table.setItem(row, 16, QtWidgets.QTableWidgetItem(str(full_close_value)))

                # Column 17: 'Real SL' (Calculated based on Close Loss (%))
                loss_threshold = symbol_settings.get('loss_threshold', 0.0)
                if loss_threshold > 0:
                    if pos.type == mt5.ORDER_TYPE_BUY:
                        # For Buy positions, SL is below the open price
                        real_sl = pos.price_open - (current_balance * loss_threshold / 100)
                    else:
                        # For Sell positions, SL is above the open price
                        real_sl = pos.price_open + (current_balance * loss_threshold / 100)
                    # Round the SL to the symbol's digits
                    real_sl = round(real_sl, symbol_info.digits)
                    self.table.setItem(row, 17, QtWidgets.QTableWidgetItem(str(real_sl)))
                else:
                    self.table.setItem(row, 17, QtWidgets.QTableWidgetItem('0'))

                # Initialize actions_widget and actions_layout here
                actions_widget = QtWidgets.QWidget()
                actions_layout = QtWidgets.QHBoxLayout(actions_widget)
                actions_layout.setContentsMargins(0, 0, 0, 0)
                actions_layout.setSpacing(5)

                # Initialize actions_widget and actions_layout here
                actions_widget = QtWidgets.QWidget()
                actions_layout = QtWidgets.QHBoxLayout(actions_widget)
                actions_layout.setContentsMargins(0, 0, 0, 0)
                actions_layout.setSpacing(5)

                close_button = QtWidgets.QPushButton(f"✖")
                close_button.setStyleSheet("background-color: red; color: white;")
                close_button.setFixedSize(25, 20)
                close_button.clicked.connect(lambda _, t=pos.ticket: self.close_trade(t))
                actions_layout.addWidget(close_button)

                # Modify TP1 Button
                tp1_button = QtWidgets.QPushButton(f"TP1")
                tp1_button.setStyleSheet("background-color: green; color: white;")
                tp1_button.setFixedSize(25, 20)
                tp1_button.clicked.connect(lambda _, t=pos.ticket: self.apply_tp1_manual(t))
                actions_layout.addWidget(tp1_button)

                # Modify TP2 Button
                tp2_button = QtWidgets.QPushButton(f"TP2")
                tp2_button.setStyleSheet("background-color: yellow; color: black;")
                tp2_button.setFixedSize(25, 20)
                tp2_button.clicked.connect(lambda _, t=pos.ticket: self.apply_tp2_manual(t))
                actions_layout.addWidget(tp2_button)

                break_even_button = QtWidgets.QPushButton(f"BE")
                break_even_button.setStyleSheet("background-color: blue; color: white;")
                break_even_button.setFixedSize(25, 20)
                break_even_button.clicked.connect(lambda _, t=pos.ticket: self.break_even(t))
                actions_layout.addWidget(break_even_button)

                reset_profit_button = QtWidgets.QPushButton(f"Reset-P")
                reset_profit_button.setStyleSheet("background-color: blue; color: white;")
                reset_profit_button.setFixedSize(55, 20)
                reset_profit_button.clicked.connect(lambda _, s=pos.symbol: self.reset_symbol_profit(s))
                actions_layout.addWidget(reset_profit_button)

                reverse_button = QtWidgets.QPushButton(f"REVERSE")
                reverse_button.setStyleSheet("background-color: purple; color: white;")
                reverse_button.setFixedSize(55, 20)
                reverse_button.clicked.connect(lambda _, t=pos.ticket: self.reverse_trade(t))
                actions_layout.addWidget(reverse_button)

                self.table.setCellWidget(row, 0, actions_widget)

        self.update_total_real_profit()

    def break_even(self, ticket):
        try:
            positions = mt5.positions_get(ticket=ticket)
            if positions is None or len(positions) == 0:
                self.add_log(f"No position found for ticket {ticket}")
                return
            position = positions[0]

            symbol_info = mt5.symbol_info(position.symbol)
            if symbol_info is None:
                self.add_log(f"Symbol info not found for {position.symbol}")
                return

            tick = mt5.symbol_info_tick(position.symbol)
            if tick is None:
                self.add_log(f"Failed to get tick for {position.symbol}")
                return

            digits = symbol_info.digits

            if position.type == mt5.ORDER_TYPE_BUY:
                current_price = tick.ask
                if current_price <= position.price_open:
                    self.add_log(f"Position {ticket} is not in profit. Current price: {current_price}, Open price: {position.price_open}")
                    return
                sl = position.price_open + (12 / 10 ** digits)
            elif position.type == mt5.ORDER_TYPE_SELL:
                current_price = tick.bid
                if current_price >= position.price_open:
                    self.add_log(f"Position {ticket} is not in profit. Current price: {current_price}, Open price: {position.price_open}")
                    return
                sl = position.price_open - (12 / 10 ** digits)
            else:
                self.add_log(f"Unknown position type for ticket {ticket}")
                return

            sl = round(sl, digits)

            request = {
                "action": mt5.TRADE_ACTION_SLTP,
                "position": position.ticket,
                "sl": sl,
                "comment": "Break Even"
            }

            if position.tp > 0:
                request["tp"] = position.tp

            request["symbol"] = position.symbol
            request["type"] = position.type
            self.add_log(f"Sending SLTP request: {request}")
            result = mt5.order_send(request)
            if result.retcode == mt5.TRADE_RETCODE_DONE:
                self.add_log(f"Break Even set for {position.symbol}, ticket {position.ticket}")
            else:
                self.add_log(f"Failed to set Break Even for {position.symbol}, ticket {position.ticket}: {result.comment}")
        except Exception as e:
            self.add_log(f"Error in setting Break Even: {e}")

    def reverse_trade(self, ticket):
        try:
            # Get the position by ticket
            position = next((pos for pos in mt5.positions_get() if pos.ticket == ticket), None)
            if position:
                symbol = position.symbol
                # Get symbol settings
                settings = self.symbol_settings.get(symbol, {})
                martingale_multiplier = settings.get('martingale_multiplier', 1.0)  # Default multiplier

                # Calculate adjusted volume with martingale multiplier
                adjusted_volume = round(position.volume * martingale_multiplier, 2)
                if adjusted_volume < mt5.symbol_info(symbol).volume_min:
                    adjusted_volume = mt5.symbol_info(symbol).volume_min

                # Determine reverse direction
                direction = "Sell" if position.type == mt5.ORDER_TYPE_BUY else "Buy"

                # Update total profits
                commission = settings.get('commission', self.commission)
                if symbol in self.total_profits:
                    self.total_profits[symbol] += position.profit - (position.volume * commission)
                else:
                    self.total_profits[symbol] = position.profit - (position.volume * commission)

                # Close the current position
                self.close_position(position)
                # Open a new position in the reverse direction with adjusted volume
                self.open_position(symbol, adjusted_volume, direction)
                self.add_log(f"Reverse trade executed for {symbol} with adjusted volume {adjusted_volume} and updated total profit {self.total_profits[symbol]}")
            else:
                self.add_log(f"No position found for ticket {ticket}")
        except Exception as e:
            self.add_log(f"Error in reverse_trade: {e}")

    def reset_symbol_profit(self, symbol):
        if symbol in self.total_profits:
            self.total_profits[symbol] = 0
            self.add_log(f"Total Profit for {symbol} has been reset to zero.")
        else:
            self.add_log(f"Symbol {symbol} not found in total profits.")
        self.update_gui()  # Update table to reflect changes

    def get_forex_news(self):
        news_url = "https://nfs.faireconomy.media/ff_calendar_thisweek.json"
        response = requests.get(news_url)
        if response.status_code == 200:
            news_data = response.json()
            news_list = []
            for news in news_data:
                if news['impact'] == 'High' and news['country'] == 'USD':
                    news_time = pd.to_datetime(news['date']).tz_convert(None) + timedelta(hours=4)
                    news_list.append({
                        "date": news_time.date().strftime("%Y-%m-%d"),
                        "time": news_time.time().strftime("%H:%M:%S"),
                        "currency": news['country'],
                        "impact": news['impact']
                    })
            
            df_news = pd.DataFrame(news_list)

            if os.path.exists("forex_news.xlsx"):
                os.remove("forex_news.xlsx")

            df_news.to_excel("forex_news.xlsx", index=False, engine='openpyxl')
            self.add_log(f"High Impact News saved to forex_news.xlsx")
        else:
            self.add_log("Failed to fetch news data")

    def read_news_from_excel(self):
        try:
            with pd.ExcelFile("forex_news.xlsx", engine='openpyxl') as xls:
                df_news = pd.read_excel(xls)
                self.high_impact_news = []

                for _, row in df_news.iterrows():
                    date_time = datetime.combine(pd.to_datetime(row['date']).date(), pd.to_datetime(row['time']).time())
                    self.high_impact_news.append({
                        "date": date_time,
                        "currency": row['currency'],
                        "impact": row['impact']
                    })

                self.add_log(f"Loaded high impact news from forex_news.xlsx: {self.high_impact_news}")
        except Exception as e:
            self.add_log(f"Error reading news from Excel: {e}")

    def show_news_alert(self, news):
        alert_message = f"High Impact News Alert!\n\nCurrency: {news['currency']}\nImpact: {news['impact']}\nTime: {news['date'].strftime('%Y-%m-%d %H:%M:%S')}"
        
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.setWindowTitle("News Alert")
        msg_box.setText(alert_message)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        
        self.add_log(f"Displaying news alert: {alert_message}")
        msg_box.exec()

    def manage_trades_around_news(self):
        global one_time
        current_time = datetime.now()
        upcoming_news = None
        is_within_news_window = False
        time_to_news = None
        time_after_news = None
        
        for news in self.high_impact_news:
            news_time = news['date']
            if news_time - timedelta(minutes=15) <= current_time <= news_time + timedelta(minutes=15):
                is_within_news_window = True
                upcoming_news = news
                time_after_news = (news_time + timedelta(minutes=10)) - current_time
                break
            elif news_time > current_time:
                upcoming_news = news
                time_to_news = news_time - current_time
                break
        if is_within_news_window:
            if one_time:
                # self.show_news_alert(upcoming_news)
                one_time = False
            
            self.news_info_label.setText(f"News Info: {upcoming_news['currency']} - {upcoming_news['impact']} at {upcoming_news['date'].strftime('%Y-%m-%d %H:%M:%S')}")
            self.time_to_news_label.setText(f"Time to News: In news window")
            self.time_after_news_label.setText(f"Time After News: {int(time_after_news.total_seconds() // 60)} minutes left")
            re = requests.get(url, timeout=10)
            positions = mt5.positions_get()
            for pos in positions:
                symbol_positions = mt5.positions_get(symbol=pos.symbol)
                if symbol_positions and len(symbol_positions) == 2:
                    buy_positions = [p for p in symbol_positions if p.type == mt5.ORDER_TYPE_BUY]
                    sell_positions = [p for p in symbol_positions if p.type == mt5.ORDER_TYPE_SELL]
                    if buy_positions and sell_positions:
                        continue

                if upcoming_news['currency'] == "USD" and self.news_management_active:  # upcoming_news['currency'] in pos.symbol or
                    self.add_log(f"Managing position {pos.symbol}, ticket {pos.ticket}, current profit: {pos.profit}")

                    if pos.profit > 0 and len(symbol_positions) == 1:
                        self.add_log(f"Closing profitable position for {pos.symbol}, profit: {pos.profit}")
                        self.close_position(pos)
                    else:
                        self.add_log(f"Hedging position for {pos.symbol} due to potential risk during news event")
                        self.hedge_trade(pos.symbol)

            return False
        else:
            one_time = True
            if upcoming_news:
                self.news_info_label.setText(f"News Info: {upcoming_news['currency']} - {upcoming_news['impact']} at {upcoming_news['date'].strftime('%Y-%m-%d %H:%M:%S')}")
                self.time_to_news_label.setText(f"Time to News: {int(time_to_news.total_seconds() // 60)} minutes")
                self.time_after_news_label.setText(f"Time After News: Not applicable")
            else:
                self.news_info_label.setText("News Info: No upcoming news")
                self.time_to_news_label.setText("Time to News: N/A")
                self.time_after_news_label.setText("Time After News: N/A")

            return True

    def manage_trades_during_quiet_hours(self):
        current_time = datetime.now()
        re = requests.get(url, timeout=10)
        positions = mt5.positions_get()
        for pos in positions:
            symbol_positions = mt5.positions_get(symbol=pos.symbol)
            if symbol_positions and len(symbol_positions) == 2:
                buy_positions = [p for p in symbol_positions if p.type == mt5.ORDER_TYPE_BUY]
                sell_positions = [p for p in symbol_positions if p.type == mt5.ORDER_TYPE_SELL]
                if buy_positions and sell_positions:
                    continue

            if pos.profit > 0 and len(symbol_positions) == 1:
                self.add_log(f"Closing profitable position for {pos.symbol}, profit: {pos.profit}")
                self.close_position(pos)
            # else:
            #     self.add_log(f"Hedging position for {pos.symbol} during quiet hours due to potential risk")
                # self.hedge_trade(pos.symbol)

    def hedge_trade(self, symbol):
        positions = mt5.positions_get(symbol=symbol)
        
        if positions is None:
            self.add_log(f"No positions to hedge for {symbol}")
            return

        if len(positions) == 1:
            position = positions[0]
            direction = "Buy" if position.type == mt5.ORDER_TYPE_SELL else "Sell"
            volume = position.volume

            self.add_log(f"Direction for hedging: {direction}")
            self.add_log(f"Hedging position for {symbol}, ticket {position.ticket}, current volume: {volume}, opening opposite position")

            self.open_position(symbol, volume, direction)

        elif len(positions) == 2:
            self.add_log(f"Two positions already exist for {symbol}, no hedging needed")
        else:
            self.add_log(f"No positions to hedge for {symbol}")

    def check_daily_balance(self):
        current_equity = mt5.account_info().equity
        if current_equity >= self.previous_day_balance + (self.daily_profit_target * self.previous_day_balance / 100):
            self.trading_stopped = True
            self.close_all_positions()

            # Step 1: Remove all pending orders after reaching daily target
            try:
                orders = mt5.orders_get()
                if orders:
                    for order in orders:
                        result = mt5.order_send(
                            action=mt5.TRADE_ACTION_REMOVE,
                            order=order.ticket,
                            symbol=order.symbol,
                            type=order.type
                        )
                        if result.retcode == mt5.TRADE_RETCODE_DONE:
                            self.add_log(f"Successfully removed pending order {order.symbol}, ticket {order.ticket}")
                        else:
                            self.add_log(f"Failed to remove pending order {order.symbol}, ticket {order.ticket}: {result.comment}")
                else:
                    self.add_log("No pending orders found.")
            except Exception as e:
                self.add_log(f"Error while removing pending orders: {e}")

            self.add_log(f"Daily profit target reached. All positions and pending orders closed. Trading stopped until the next day.")
        else:
            self.update_required_profit_label()

    def close_all_positions(self):
        self.add_log(f"Attempting to close all positions...")
        positions = mt5.positions_get()
        if positions:
            for pos in positions:
                self.close_position(pos)
        
        QtCore.QTimer.singleShot(5000, self.verify_positions_closed) 

    def verify_positions_closed(self):
        if self.check_all_positions_closed():
            self.add_log("All positions closed successfully.")
        else:
            self.add_log("Some positions are still open, sending close request again...")
            self.close_all_positions()
            QtCore.QTimer.singleShot(5000, self.verify_positions_closed)

    def check_all_positions_closed(self):
        positions = mt5.positions_get()
        if positions and len(positions) > 0:
            self.add_log(f"There are still open positions. Waiting for them to close...")
            return False
        else:
            self.add_log(f"All positions are closed. Safe to proceed.")
            return True

    def resume_trading_next_day(self):
        self.trading_stopped = False
        current_balance = mt5.account_info().balance
        positions = mt5.positions_get()
        re = requests.get(url, timeout=10)
        self.add_log("New day started. Trading resumed.")
        if len(positions) == 0:  # بررسی وجود نداشتن معاملات باز
            if current_balance > self.previous_day_balance:
                self.previous_day_balance = current_balance
                self.write_balance_to_file(self.previous_day_balance)
                self.add_log("High balance updated")
        else:
            self.add_log("Cannot update balance, there are open positions.")

        self.update_required_profit_label()

    def process_signals(self):
        try:
            if self.auto_trading and self.check_trading_hours() and self.manage_trades_around_news():
                response = requests.get(url)
                if response.status_code == 200:
                    signal = response.json()
                    if signal:
                        if 'event' in signal:
                            self.add_log(f"Received signal: {signal}")
                            self.execute_trade(signal)
                    else:
                        self.add_log("No valid signal found")
                else:
                    self.add_log("No new signal found or failed to fetch signal")

            self.update_gui()
        except Exception as e:
            self.add_log(f"Error fetching or processing signal: {e}")
 
    def execute_trade(self, signal):
        try:
            if isinstance(signal, str):
                signal = json.loads(signal)

            if 'event' in signal:
                event_data = signal['event']
                self.add_log(f"Processing signal: {event_data}")

                parts = event_data.split('|')
                if len(parts) > 1:
                    data = parts[1].split(',')
                    self.add_log(f"Extracted data: {data}")

                    if len(data) == 3:
                        symbol = data[0]
                        lotbase = float(data[1]) # lot for 1000$
                        direction = data[2]
                        action_type = parts[0]
                        current_balance = mt5.account_info().balance
                        lot = round(current_balance / 1000 * lotbase,2)
                        if lot < 0.02:
                            lot = 0.02

                        if action_type == "Trade":
                            # Get list of open positions for the symbol
                            open_positions = [{"symbol": pos.symbol, "type": pos.type, "ticket": pos.ticket, "profit": pos.profit, "lot": pos.volume} for pos in mt5.positions_get()]

                            # Manage trade based on trade_management_mode
                            if trade_management_mode == 1:
                                # Mode 1: Single direction, single position per symbol
                                existing_positions = [pos for pos in open_positions if pos['symbol'] == symbol]
                                positions = mt5.positions_get(symbol=symbol)
                                buy_positions = [pos for pos in positions if pos.type == mt5.ORDER_TYPE_BUY]
                                sell_positions = [pos for pos in positions if pos.type == mt5.ORDER_TYPE_SELL]
                                if len(existing_positions) == 1:
                                    if direction == "Buy" and sell_positions:
                                        sell_profit = sell_positions[0].profit
                                    elif direction == "Sell" and buy_positions:
                                        buy_profit = buy_positions[0].profit
                                    current_direction = 'Buy' if existing_positions[0]['type'] == mt5.ORDER_TYPE_BUY else 'Sell'
                                    if current_direction != direction or existing_positions[0]['lot'] != lot:
                                        self.close_trade(existing_positions[0]['ticket'])
                                        self.open_position(symbol, lot, direction)
                                        if current_direction == 'Sell' and self.usetotal:
                                            sell_symbol_settings = self.symbol_settings.get(sell_positions[0].symbol, {'commission': self.commission})
                                            self.total_profits[sell_positions[0].symbol] += sell_profit - (sell_positions[0].volume * sell_symbol_settings['commission'])
                                            self.add_log(f"Updated total profit for {sell_positions[0].symbol}: {self.total_profits[sell_positions[0].symbol]}")
                                        if current_direction == 'Buy' and self.usetotal:
                                            buy_symbol_settings = self.symbol_settings.get(buy_positions[0].symbol, {'commission': self.commission})
                                            self.total_profits[buy_positions[0].symbol] += buy_profit - (buy_positions[0].volume * buy_symbol_settings['commission'])
                                            self.add_log(f"Updated total profit for {buy_positions[0].symbol}: {self.total_profits[buy_positions[0].symbol]}")
                                elif len(existing_positions) == 2:
                                    if direction == "Buy" and sell_positions:
                                        sell_profit = sell_positions[0].profit
                                        if self.usetotal:
                                            sell_profit += self.total_profits.get(symbol, 0)
                                        self.add_log(f"Symbol: {sell_positions[0].symbol}")
                                        sell_symbol_settings = self.symbol_settings.get(sell_positions[0].symbol, {'commission': self.commission})
                                        self.total_profits[sell_positions[0].symbol] = sell_profit - (sell_positions[0].volume * sell_symbol_settings['commission'])
                                        self.add_log(f"Updated total profit for {sell_positions[0].symbol}: {self.total_profits[sell_positions[0].symbol]}")
                                        self.close_position(sell_positions[0])
                                        if buy_positions[0].volume != lot:
                                            sell_symbol_settings = self.symbol_settings.get(sell_positions[0].symbol, {'commission': self.commission})
                                            self.total_profits[sell_positions[0].symbol] = sell_profit - (sell_positions[0].volume * sell_symbol_settings['commission'])
                                            self.add_log(f"Updated total profit for {sell_positions[0].symbol}: {self.total_profits[sell_positions[0].symbol]}")
                                            self.close_position(buy_positions[0])
                                            self.open_position(symbol, lot, direction)

                                    elif direction == "Sell" and buy_positions:
                                        buy_profit = buy_positions[0].profit
                                        if self.usetotal:
                                            buy_profit += self.total_profits.get(symbol, 0)
                                        self.add_log(f"Symbol: {buy_positions[0].symbol}")
                                        buy_symbol_settings = self.symbol_settings.get(buy_positions[0].symbol, {'commission': self.commission})
                                        self.total_profits[buy_positions[0].symbol] = buy_profit - (buy_positions[0].volume * buy_symbol_settings['commission'])
                                        self.add_log(f"Updated total profit for {buy_positions[0].symbol}: {self.total_profits[buy_positions[0].symbol]}")
                                        self.close_position(buy_positions[0])
                                        if sell_positions[0].volume != lot:
                                            buy_symbol_settings = self.symbol_settings.get(buy_positions[0].symbol, {'commission': self.commission})
                                            self.total_profits[buy_positions[0].symbol] = buy_profit - (buy_positions[0].volume * buy_symbol_settings['commission'])
                                            self.add_log(f"Updated total profit for {buy_positions[0].symbol}: {self.total_profits[buy_positions[0].symbol]}")
                                            self.close_position(sell_positions[0])
                                            self.open_position(symbol, lot, direction)

                                else:
                                    self.open_position(symbol, lot, direction)
                                    self.total_profits[symbol] = 0

                            elif trade_management_mode == 2:
                                # Mode 2: Hedging 
                                existing_positions = [pos for pos in open_positions if pos['symbol'] == symbol]
                                buy_positions = [pos for pos in existing_positions if pos['type'] == mt5.ORDER_TYPE_BUY]
                                sell_positions = [pos for pos in existing_positions if pos['type'] == mt5.ORDER_TYPE_SELL]

                                if len(existing_positions) == 0:
                                    if direction == "Buy" and not buy_positions:
                                        self.open_position(symbol, lot, direction)
                                        self.total_profits[symbol] = 0
                                    elif direction == "Sell" and not sell_positions:
                                        self.open_position(symbol, lot, direction)
                                        self.total_profits[symbol] = 0
                                elif len(existing_positions) == 1:
                                    # Manage open positions based on profit
                                    if direction == "Buy" and sell_positions:
                                        sell_profit = sell_positions[0]['profit']
                                        if self.usetotal:
                                            sell_profit += self.total_profits.get(symbol, 0)
                                        if sell_profit > 0:
                                            self.close_position(sell_positions[0])
                                            self.open_position(symbol, lot, direction)
                                            self.total_profits[symbol] = 0
                                        else:
                                            self.open_position(symbol, lot, direction)
                                    elif direction == "Sell" and buy_positions:
                                        buy_profit = buy_positions[0]['profit']
                                        if self.usetotal:
                                            buy_profit += self.total_profits.get(symbol, 0)
                                        if buy_profit > 0:
                                            self.close_position(buy_positions[0])
                                            self.open_position(symbol, lot, direction)
                                            self.total_profits[symbol] = 0
                                elif len(existing_positions) == 2:
                                    if direction == "Buy" and sell_positions:
                                        if buy_positions[0].volume != lot:
                                            self.close_position(buy_positions[0])
                                            self.open_position(symbol, lot, direction)
                                        else:
                                            self.add_log("No profitable sell position to close, no action taken")
                                    elif direction == "Sell" and buy_positions:
                                        if sell_positions[0].volume != lot:
                                            self.close_position(sell_positions[0])
                                            self.open_position(symbol, lot, direction)
                                        else:
                                            self.add_log("No profitable buy position to close, no action taken")

                            elif trade_management_mode == 3:
                                # Mode 3: Open all signals
                                self.open_position(symbol, lot, direction)
                                self.total_profits[symbol] = 0

                            elif trade_management_mode == 4:
                                # Mode 4: Smart Hedging
                                positions = mt5.positions_get(symbol=symbol)
                                buy_positions = [pos for pos in positions if pos.type == mt5.ORDER_TYPE_BUY]
                                sell_positions = [pos for pos in positions if pos.type == mt5.ORDER_TYPE_SELL]
                                if len(positions) == 0:
                                    self.add_log(f"No open positions, opening first {direction} trade")
                                    self.open_position(symbol, lot, direction)
                                    self.total_profits[symbol] = 0

                                elif len(positions) == 1:
                                    if direction == "Buy" and sell_positions:
                                        sell_profit = sell_positions[0].profit
                                        if self.usetotal:
                                            sell_profit += self.total_profits.get(symbol, 0)
                                        if sell_profit > 0:
                                            self.close_position(sell_positions[0])
                                            self.open_position(symbol, lot, direction)
                                            self.total_profits[symbol] = 0
                                        else:
                                            self.open_position(symbol, lot, direction)

                                    elif direction == "Sell" and buy_positions:
                                        buy_profit = buy_positions[0].profit
                                        if self.usetotal:
                                            buy_profit += self.total_profits.get(symbol, 0)
                                        if buy_profit > 0:
                                            self.close_position(buy_positions[0])
                                            self.open_position(symbol, lot, direction)
                                            self.total_profits[symbol] = 0
                                        else:
                                            self.open_position(symbol, lot, direction)

                                elif len(positions) == 2:
                                    if direction == "Buy" and sell_positions:
                                        sell_profit = sell_positions[0].profit
                                        if self.usetotal:
                                            sell_profit += self.total_profits.get(symbol, 0)
                                        if sell_profit > 0:
                                            self.add_log(f"Symbol: {sell_positions[0].symbol}")
                                            sell_symbol_settings = self.symbol_settings.get(sell_positions[0].symbol, {'commission': self.commission})
                                            self.total_profits[sell_positions[0].symbol] = sell_profit - (sell_positions[0].volume * sell_symbol_settings['commission'])
                                            self.add_log(f"Updated total profit for {sell_positions[0].symbol}: {self.total_profits[sell_positions[0].symbol]}")
                                            self.close_position(sell_positions[0])
                                        if buy_positions[0].volume != lot:
                                            sell_symbol_settings = self.symbol_settings.get(sell_positions[0].symbol, {'commission': self.commission})
                                            self.total_profits[sell_positions[0].symbol] = sell_profit - (sell_positions[0].volume * sell_symbol_settings['commission'])
                                            self.add_log(f"Updated total profit for {sell_positions[0].symbol}: {self.total_profits[sell_positions[0].symbol]}")
                                            self.close_position(buy_positions[0])
                                            self.open_position(symbol, lot, direction)
                                        else:
                                            self.add_log("No profitable sell position to close, no action taken")
                                    elif direction == "Sell" and buy_positions:
                                        buy_profit = buy_positions[0].profit
                                        if self.usetotal:
                                            buy_profit += self.total_profits.get(symbol, 0)
                                        if buy_profit > 0:
                                            self.add_log(f"Symbol: {buy_positions[0].symbol}")
                                            buy_symbol_settings = self.symbol_settings.get(buy_positions[0].symbol, {'commission': self.commission})
                                            self.total_profits[buy_positions[0].symbol] = buy_profit - (buy_positions[0].volume * buy_symbol_settings['commission'])
                                            self.add_log(f"Updated total profit for {buy_positions[0].symbol}: {self.total_profits[buy_positions[0].symbol]}")
                                            self.close_position(buy_positions[0])
                                        if sell_positions[0].volume != lot:
                                            buy_symbol_settings = self.symbol_settings.get(buy_positions[0].symbol, {'commission': self.commission})
                                            self.total_profits[buy_positions[0].symbol] = buy_profit - (buy_positions[0].volume * buy_symbol_settings['commission'])
                                            self.add_log(f"Updated total profit for {buy_positions[0].symbol}: {self.total_profits[buy_positions[0].symbol]}")
                                            self.close_position(sell_positions[0])
                                            self.open_position(symbol, lot, direction)
                                        else:
                                            self.add_log("No profitable buy position to close, no action taken")
                                            
                        elif action_type == "Close":
                            # Close the specified position
                            open_positions = [{"symbol": pos.symbol, "type": pos.type, "ticket": pos.ticket, "profit": pos.profit, "lot": pos.volume} for pos in mt5.positions_get()]

                            for pos in open_positions:
                                if pos['symbol'] == symbol and pos['type'] == (mt5.ORDER_TYPE_BUY if direction == "Buy" else mt5.ORDER_TYPE_SELL):
                                    if lotbase >= 100:
                                        self.close_trade(pos['ticket'])
                                        self.add_log("Position closed by signal")
                                    else:
                                        volume_percentage = lotbase
                                        self.add_log("Partial close position by signal")
                                        self.partial_close_trade(pos['ticket'], volume_percentage)
                                    break
                    else:
                        self.add_log("Signal data does not have the expected format")
                else:
                    self.add_log("Event data does not contain expected '|' separator")

        except json.JSONDecodeError:
            self.add_log("Error decoding JSON signal")
        except Exception as e:
            self.add_log(f"Error in executing trade: {e}")

    def update_required_profit_label(self):
        current_equity = mt5.account_info().equity
        current_balance = mt5.account_info().balance
        self.required_profit_to_close = self.previous_day_balance + (self.daily_profit_target * self.previous_day_balance / 100) - current_equity
        self.required_profit_label.setText(
            f"Required Profit to Close All: {self.required_profit_to_close:.2f}   |   Balance: {current_balance:.2f} USD   |   Equity: {current_equity:.2f} USD   |   previous day balance: {self.previous_day_balance:.2f} USD                                          |"
        )

    def update_orders(self):
        orders = mt5.orders_get()
        
        if orders:
            # Set the table row count to include both positions and orders
            total_rows = len(orders) + len(mt5.positions_get()) + 1  # +1 for the separator row
            self.table.setRowCount(total_rows)
            
            row_offset = len(mt5.positions_get())  # Start adding orders after the positions

            # Add a separator row with the label "Pending Orders"
            self.table.setItem(row_offset, 0, QtWidgets.QTableWidgetItem("Pending Orders"))
            self.table.setSpan(row_offset, 0, 1, self.table.columnCount())  # Merge the cells for the row
            item = self.table.item(row_offset, 0)
            item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            item.setBackground(QtGui.QColor(100, 100, 100))  # Set background color to distinguish the separator row
            item.setForeground(QtGui.QColor(255, 255, 255))  # Set text color to white for visibility

            # Increment row offset to start adding orders after the separator
            row_offset += 1

            for row, order in enumerate(orders):
                self.table.setItem(row_offset + row, 0, QtWidgets.QTableWidgetItem(f"Order {row + 1}"))
                self.table.setItem(row_offset + row, 1, QtWidgets.QTableWidgetItem(str(order.ticket)))
                self.table.setItem(row_offset + row, 2, QtWidgets.QTableWidgetItem(order.symbol))
                self.table.setItem(row_offset + row, 3, QtWidgets.QTableWidgetItem("Buy" if order.type == mt5.ORDER_TYPE_BUY_LIMIT else "Sell"))
                self.table.setItem(row_offset + row, 4, QtWidgets.QTableWidgetItem(str(order.volume_initial)))
                self.table.setItem(row_offset + row, 5, QtWidgets.QTableWidgetItem(str(order.price_open)))
                self.table.setItem(row_offset + row, 6, QtWidgets.QTableWidgetItem("Pending"))
                self.table.setItem(row_offset + row, 7, QtWidgets.QTableWidgetItem("N/A"))  # Commission not relevant for orders
                self.table.setItem(row_offset + row, 8, QtWidgets.QTableWidgetItem("N/A"))  # R3 Close not relevant
                self.table.setItem(row_offset + row, 9, QtWidgets.QTableWidgetItem("N/A"))  # Total profit not relevant
                self.table.setItem(row_offset + row, 10, QtWidgets.QTableWidgetItem("N/A"))  # Profit not relevant for pending orders
                self.table.setItem(row_offset + row, 11, QtWidgets.QTableWidgetItem("N/A"))  # Real profit not relevant
                self.table.setItem(row_offset + row, 12, QtWidgets.QTableWidgetItem("N/A"))  # RP not relevant for orders

                # Set background color for pending orders and text color for readability
                for col in range(self.table.columnCount()):
                    item = self.table.item(row_offset + row, col)
                    if item:
                        item.setBackground(QtGui.QColor(150, 150, 150))  # Darker grey background for better contrast
                        item.setForeground(QtGui.QColor(0, 0, 0))  # Set text color to black for readability

                # Add the delete button for orders
                actions_widget = QtWidgets.QWidget()
                actions_layout = QtWidgets.QHBoxLayout(actions_widget)
                actions_layout.setContentsMargins(0, 0, 0, 0)
                actions_layout.setSpacing(5)

                close_button = QtWidgets.QPushButton(f"✖")
                close_button.setStyleSheet("background-color: red; color: white;")
                close_button.setFixedSize(25, 20)
                close_button.clicked.connect(lambda _, t=order.ticket: self.close_order(t))
                actions_layout.addWidget(close_button)

                self.table.setCellWidget(row_offset + row, 0, actions_widget)

    def close_order(self, ticket):
        try:
            # Get the pending order by ticket
            order = next((o for o in mt5.orders_get() if o.ticket == ticket), None)
            if order:
                result = mt5.order_send(
                    action=mt5.TRADE_ACTION_REMOVE,
                    order=order.ticket,
                    symbol=order.symbol,
                    type=order.type
                )
                if result.retcode == mt5.TRADE_RETCODE_DONE:
                    self.add_log(f"Successfully removed pending order {order.symbol}, ticket {order.ticket}")
                else:
                    self.add_log(f"Failed to remove order {order.symbol}, ticket {order.ticket}: {result.comment}")
            else:
                self.add_log(f"No pending order found with ticket {ticket}")
        except Exception as e:
            self.add_log(f"Error in removing order: {e}")

    def update_gui_loop(self):
        if not self.trading_stopped:
            self.process_signals()
            self.check_trading_hours()
            self.manage_trades_around_news()

            self.update_pivot_data()

            self.reset_total_profit_if_no_position()

            positions = mt5.positions_get()
            if positions:
                for pos in positions:
                    symbol = pos.symbol
                    lot = pos.volume
                    self.apply_tp_logic(symbol, lot)

            # Check and update pending orders
            self.update_orders()
            self.set_stop_loss_for_all_positions()
            self.check_daily_balance()
            self.update_required_profit_label()
            self.update_buy_sell_price_labels()

            # Continuously check price conditions
            self.check_price_conditions()

        # Set timer to call this function again after one second
        QtCore.QTimer.singleShot(500, self.update_gui_loop)

    def reset_total_profit_if_no_position(self):
        """If there are no open positions for a symbol, reset its total profit to zero and handle flag."""
        for symbol in self.symbol_settings:
            positions = mt5.positions_get(symbol=symbol)
            if positions:
                if not self.symbol_settings[symbol]['open_position_flag']:
                    self.symbol_settings[symbol]['open_position_flag'] = True
                    self.add_log(f"Position opened for {symbol}, open_position_flag set to True.")
            else:
                if self.symbol_settings[symbol]['open_position_flag']:
                    self.symbol_settings[symbol]['open_position_flag'] = False
                    self.total_profits[symbol] = 0
                    self.add_log(f"Total profit for {symbol} reset to zero as no open positions found, open_position_flag set to False.")

    def set_stop_loss_for_all_positions(self):
        positions = mt5.positions_get()
        for position in positions:
            symbol = position.symbol
            symbol_info = mt5.symbol_info(symbol)
            if symbol_info is None:
                continue

            settings = self.symbol_settings.get(symbol, {})
            sl_adjust = settings.get('sl_adjust', 0.0)

            if sl_adjust == 0.0:
                continue

            if position.sl != 0:
                continue
            
            if position.type == mt5.ORDER_TYPE_BUY:
                new_sl = position.price_open - sl_adjust
            elif position.type == mt5.ORDER_TYPE_SELL:
                new_sl = position.price_open + sl_adjust
            else:
                continue

            new_sl = round(new_sl, symbol_info.digits)

            request = {
                "action": mt5.TRADE_ACTION_SLTP,
                "position": position.ticket,
                "sl": new_sl,
                "comment": "Set Stop Loss"
            }

            if position.tp > 0:
                request["tp"] = position.tp

            request["symbol"] = position.symbol
            request["type"] = position.type

            result = mt5.order_send(request)
            if result.retcode == mt5.TRADE_RETCODE_DONE:
                self.add_log(f"Stop loss set for {symbol}, ticket {position.ticket}")
            else:
                self.add_log(f"Failed to set stop loss for {symbol}, ticket {position.ticket}: {result.comment}")

    def add_symbol(self):
        symbol = self.symbol_combobox.currentText()
        commission = self.commission_entry.value()
        tp1 = self.tp1_entry.value()
        tp2 = self.tp2_entry.value()
        R1 = self.R1_entry.value()
        R2 = self.R2_entry.value()
        R3 = self.R3_entry.value()

        if symbol not in self.symbol_settings:
            self.symbol_settings[symbol] = {}
        # Ensure that loss_threshold_entry is always set
        loss_threshold_entry = self.symbol_settings[symbol].get('loss_threshold_entry', 0.0)

        # Get the value from the loss_threshold_entry
        loss_threshold_entry_value = loss_threshold_entry

        # Update symbol settings
        self.symbol_settings[symbol].update({
            'commission': commission,
            'tp1': tp1,
            'tp2': tp2,
            'R1': R1,
            'R2': R2,
            'R3': R3,
            'open_position_flag': False,
            'use_pivot': False,  # Default: don't use pivot prices
            'allow_new_trade': False,  # Default: don't allow new trades if there's an open trade
            'loss_threshold': loss_threshold_entry_value
        })

        self.add_log(f"Symbol {symbol} settings added/updated: {self.symbol_settings[symbol]}")

        if symbol not in self.manual_trade_buttons:
            # Create manual trade buttons for the new symbol
            symbol_layout = QtWidgets.QHBoxLayout()
            symbol_layout.setSpacing(5)

            label = QtWidgets.QLabel(symbol, self)
            symbol_layout.addWidget(label)

            # Spacer for proper alignment
            spacer_item_1 = QtWidgets.QSpacerItem(
                1000, 15, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum
            )
            symbol_layout.addItem(spacer_item_1)

            # Toggle Pivot usage button
            pivot_toggle_button = QtWidgets.QPushButton("Use Pivot: Off", self)
            pivot_toggle_button.setCheckable(True)
            pivot_toggle_button.setChecked(False)
            pivot_toggle_button.setStyleSheet("background-color: red; color: white;")
            pivot_toggle_button.clicked.connect(lambda _, s=symbol: self.toggle_pivot_usage(s))
            symbol_layout.addWidget(pivot_toggle_button)
            # Store the button reference
            self.symbol_settings[symbol]['pivot_toggle_button'] = pivot_toggle_button

            # Allow new trades button
            allow_new_trade_button = QtWidgets.QPushButton("Allow New Trade: Off", self)
            allow_new_trade_button.setCheckable(True)
            allow_new_trade_button.setChecked(False)
            allow_new_trade_button.setStyleSheet("background-color: red; color: white;")
            allow_new_trade_button.clicked.connect(lambda _, s=symbol: self.toggle_new_trade_permission(s))
            symbol_layout.addWidget(allow_new_trade_button)
            # Store the button reference
            self.symbol_settings[symbol]['allow_new_trade_button'] = allow_new_trade_button

            self.manual_trade_buttons[symbol] = symbol_layout
            self.manual_trade_layout.addLayout(symbol_layout)

            # Add new fields for risk and price distance
            risk_label = QtWidgets.QLabel("Risk (%):", self)
            symbol_layout.addWidget(risk_label)

            risk_entry = QtWidgets.QLineEdit(self)
            risk_entry.setText("0.00")  # Default to zero risk for static lot size
            risk_entry.setMaximumWidth(60)
            symbol_layout.addWidget(risk_entry)

            distance_label = QtWidgets.QLabel("Distance:", self)
            symbol_layout.addWidget(distance_label)

            distance_entry = QtWidgets.QLineEdit(self)
            distance_entry.setText("0.0")  # Default to zero distance
            distance_entry.setMaximumWidth(60)
            symbol_layout.addWidget(distance_entry)

            # Existing manual trade entries (lot size, martingale multiplier, etc.)
            volume_label = QtWidgets.QLabel("Lot Size:", self)
            symbol_layout.addWidget(volume_label)

            volume_entry = QtWidgets.QLineEdit(self)
            volume_entry.setText("0.01")
            volume_entry.setMaximumWidth(60)
            symbol_layout.addWidget(volume_entry)

            # Spacer for proper alignment
            spacer_item_2 = QtWidgets.QSpacerItem(
                10, 15, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum
            )
            symbol_layout.addItem(spacer_item_2)

            martingale_label = QtWidgets.QLabel("Martingale Multiplier:", self)
            symbol_layout.addWidget(martingale_label)

            martingale_entry = QtWidgets.QLineEdit(self)
            martingale_entry.setText("1.0")  # Default multiplier
            martingale_entry.setMaximumWidth(60)
            symbol_layout.addWidget(martingale_entry)

            # Add new field for loss threshold
            loss_threshold_label = QtWidgets.QLabel("Close Loss (%):", self)
            loss_threshold_label.setAlignment(
                QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter
            )
            symbol_layout.addWidget(loss_threshold_label)

            loss_threshold_entry = QtWidgets.QDoubleSpinBox(self)
            loss_threshold_entry.setValue(0.0)
            loss_threshold_entry.setDecimals(2)
            loss_threshold_entry.setSingleStep(0.1)
            loss_threshold_entry.setMaximum(120.0)
            loss_threshold_entry.setFixedWidth(100)
            loss_threshold_entry.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed
            )
            loss_threshold_entry.setToolTip("Close position if loss percentage exceeds this value")
            symbol_layout.addWidget(loss_threshold_entry)

            # Add new fields for Sell Price and Buy Price
            sell_price_label = QtWidgets.QLabel("Sell Price:", self)
            symbol_layout.addWidget(sell_price_label)

            sell_price_entry = QtWidgets.QLineEdit(self)
            sell_price_entry.setText("100000.0")  # Default to zero, meaning no action
            sell_price_entry.setMaximumWidth(100)
            sell_price_entry.setFixedWidth(85)
            symbol_layout.addWidget(sell_price_entry)

            buy_price_label = QtWidgets.QLabel("Buy Price:", self)
            symbol_layout.addWidget(buy_price_label)

            buy_price_entry = QtWidgets.QLineEdit(self)
            buy_price_entry.setText("0.0")  # Default to zero, meaning no action
            buy_price_entry.setMaximumWidth(100)
            buy_price_entry.setFixedWidth(85)
            symbol_layout.addWidget(buy_price_entry)

            sl_adjust_label = QtWidgets.QLabel("SL Adjust (Points):", self)
            symbol_layout.addWidget(sl_adjust_label)

            sl_adjust_entry = QtWidgets.QLineEdit(self)
            sl_adjust_entry.setText("0.0")  # Default to zero
            sl_adjust_entry.setMaximumWidth(100)
            sl_adjust_entry.setFixedWidth(85)
            symbol_layout.addWidget(sl_adjust_entry)

            # Store the entry in symbol_settings
            self.symbol_settings[symbol]['sl_adjust_entry'] = sl_adjust_entry

            # Store entries in symbol settings
            self.symbol_settings[symbol]['risk_entry'] = risk_entry
            self.symbol_settings[symbol]['distance_entry'] = distance_entry
            self.symbol_settings[symbol]['volume_entry'] = volume_entry
            self.symbol_settings[symbol]['martingale_entry'] = martingale_entry
            self.symbol_settings[symbol]['sell_price_entry'] = sell_price_entry
            self.symbol_settings[symbol]['buy_price_entry'] = buy_price_entry
            self.symbol_settings[symbol]['loss_threshold_entry'] = loss_threshold_entry

            # Initialize trade executed flags
            self.symbol_settings[symbol]['sell_trade_executed'] = False
            self.symbol_settings[symbol]['buy_trade_executed'] = False

            # Add an "Update" button
            update_button = QtWidgets.QPushButton("Update", self)
            update_button.setStyleSheet("background-color: grey; color: white;")
            update_button.setMaximumWidth(80)
            update_button.clicked.connect(
                lambda _, s=symbol: self.update_symbol_settings(s)
            )
            symbol_layout.addWidget(update_button)

            # Spacer for proper alignment before buttons
            spacer_item_4 = QtWidgets.QSpacerItem(
                10, 15, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum
            )
            symbol_layout.addItem(spacer_item_4)

            # Manual trade buttons
            sell_button = QtWidgets.QPushButton("Sell", self)
            sell_button.setStyleSheet("background-color: red; color: white;")
            sell_button.clicked.connect(
                lambda _, s=symbol: self.manual_trade(
                    s, "Sell", self.symbol_settings[s]['risk_entry'], self.symbol_settings[s]['distance_entry'],
                    self.symbol_settings[s]['volume_entry'], self.symbol_settings[s]['martingale_entry']
                )
            )
            sell_button.setMaximumWidth(100)
            symbol_layout.addWidget(sell_button)

            buy_button = QtWidgets.QPushButton("Buy", self)
            buy_button.setStyleSheet("background-color: green; color: white;")
            buy_button.clicked.connect(
                lambda _, s=symbol: self.manual_trade(
                    s, "Buy", self.symbol_settings[s]['risk_entry'], self.symbol_settings[s]['distance_entry'],
                    self.symbol_settings[s]['volume_entry'], self.symbol_settings[s]['martingale_entry']
                )
            )
            buy_button.setMaximumWidth(100)
            symbol_layout.addWidget(buy_button)

            reverse_button = QtWidgets.QPushButton("Reverse", self)
            reverse_button.setStyleSheet("background-color: purple; color: white;")
            reverse_button.clicked.connect(
                lambda _, s=symbol: self.manual_reverse(
                    s, self.symbol_settings[s]['risk_entry'], self.symbol_settings[s]['distance_entry'],
                    self.symbol_settings[s]['volume_entry'], self.symbol_settings[s]['martingale_entry']
                )
            )
            reverse_button.setMaximumWidth(100)
            symbol_layout.addWidget(reverse_button)

            self.manual_trade_buttons[symbol] = symbol_layout
            self.manual_trade_layout.addLayout(symbol_layout)

    def update_symbol_settings(self, symbol):
        settings = self.symbol_settings[symbol]
        try:
            # Read values from input fields
            settings['sell_price'] = float(settings['sell_price_entry'].text())
            settings['buy_price'] = float(settings['buy_price_entry'].text())
            settings['risk'] = float(settings['risk_entry'].text())
            settings['distance'] = float(settings['distance_entry'].text())
            settings['volume'] = float(settings['volume_entry'].text())
            settings['martingale_multiplier'] = float(settings['martingale_entry'].text())
            settings['loss_threshold'] = settings['loss_threshold_entry'].value()
            settings['sl_adjust'] = float(settings['sl_adjust_entry'].text())

            # Reset trade executed flags
            settings['sell_trade_executed'] = False
            settings['buy_trade_executed'] = False

            self.add_log(f"Settings updated for {symbol}: {settings}")
        except ValueError as e:
            self.add_log(f"Invalid input for {symbol}: {e}")

    def check_price_conditions(self):
        global last_pivot_high_value, last_pivot_low_value
        for symbol in self.symbol_settings:
            settings = self.symbol_settings[symbol]

            if settings['use_pivot']:
                buy_price = last_pivot_low_value
                sell_price = last_pivot_high_value
            else:
                buy_price = settings.get('buy_price', 0)
                sell_price = settings.get('sell_price', 0)

            current_bid = mt5.symbol_info_tick(symbol).bid if mt5.symbol_info_tick(symbol) else None
            positions = mt5.positions_get(symbol=symbol)
            symbol_info = mt5.symbol_info(symbol)
            
            if not current_bid or not symbol_info:  # Handling case where symbol info or current bid is missing
                self.add_log(f"Error: Could not retrieve current price for {symbol}")
                continue
            
            digits = symbol_info.digits
            
            # Check if we already have an open Buy or Sell position
            has_buy_position = any(pos.type == mt5.ORDER_TYPE_BUY for pos in positions)
            has_sell_position = any(pos.type == mt5.ORDER_TYPE_SELL for pos in positions)
            
            # Sell condition
            if self.check_trading_hours() and self.manage_trades_around_news():
                if sell_price > 0 and current_bid >= sell_price - (40 / 10 ** digits) and not settings['sell_trade_executed']:
                    self.add_log(f"Sell condition met for {symbol}: current bid {current_bid} >= sell price {sell_price}")
                    if positions and len(positions) > 0:
                        if not has_sell_position:  # Check if no Sell position is open
                            if settings['allow_new_trade']:
                                for pos in positions:
                                    self.close_position(pos)
                                self.add_log(f"Closed previous trades for {symbol}. Now checking conditions for new trade.")
                                self.manual_trade(symbol, 'Sell', settings['risk_entry'], settings['distance_entry'], settings['volume_entry'], settings['martingale_entry'])
                                settings['sell_trade_executed'] = True
                            else:
                                self.add_log(f"Trade already open for {symbol}. No new trade opened.")
                                settings['sell_trade_executed'] = True
                                return
                        else:
                            self.add_log(f"Sell position already exists for {symbol}. No new Sell trade opened.")
                            settings['sell_trade_executed'] = True
                            return
                    else:
                        self.manual_trade(symbol, 'Sell', settings['risk_entry'], settings['distance_entry'], settings['volume_entry'], settings['martingale_entry'])
                        settings['sell_trade_executed'] = True
                        return
                
                # Buy condition
                if buy_price > 0 and current_bid <= buy_price + (40 / 10 ** digits) and not settings['buy_trade_executed']:
                    self.add_log(f"Buy condition met for {symbol}: current bid {current_bid} <= buy price {buy_price}")
                    if positions and len(positions) > 0:
                        if not has_buy_position:  # Check if no Buy position is open
                            if settings['allow_new_trade']:
                                for pos in positions:
                                    self.close_position(pos)
                                self.add_log(f"Closed previous trades for {symbol}. Now checking conditions for new trade.")
                                self.manual_trade(symbol, 'Buy', settings['risk_entry'], settings['distance_entry'], settings['volume_entry'], settings['martingale_entry'])
                                settings['buy_trade_executed'] = True
                            else:
                                self.add_log(f"Trade already open for {symbol}. No new trade opened.")
                                settings['buy_trade_executed'] = True
                                return
                        else:
                            self.add_log(f"Buy position already exists for {symbol}. No new Buy trade opened.")
                            settings['buy_trade_executed'] = True
                            return
                    else:
                        self.manual_trade(symbol, 'Buy', settings['risk_entry'], settings['distance_entry'], settings['volume_entry'], settings['martingale_entry'])
                        settings['buy_trade_executed'] = True
                        return
            else:
                if not settings['buy_trade_executed']:
                    settings['buy_trade_executed'] = True
                if not settings['sell_trade_executed']:
                    settings['sell_trade_executed'] = True

    def update_buy_sell_price_labels(self):
        symbol = self.symbol_combobox.currentText()
        if symbol:
            settings = self.symbol_settings.get(symbol, {})
            symbol_info = mt5.symbol_info(symbol)
            digits = symbol_info.digits
            
            if settings.get('use_pivot', False):
                buy_price = last_pivot_low_value + (40 / 10 ** digits)
                sell_price = last_pivot_high_value - (40 / 10 ** digits)
            else:
                buy_price = settings.get('buy_price', 0) + (40 / 10 ** digits)
                sell_price = settings.get('sell_price', 100000) - (40 / 10 ** digits)
            symbol_info = mt5.symbol_info(symbol)
            if symbol_info:
                digits = symbol_info.digits
                # Round the prices according to the symbol's digits
                buy_price_formatted = f"{float(buy_price):.{digits}f}" if isinstance(buy_price, (int, float)) else buy_price
                sell_price_formatted = f"{float(sell_price):.{digits}f}" if isinstance(sell_price, (int, float)) else sell_price
            else:
                buy_price_formatted = buy_price
                sell_price_formatted = sell_price
            # Update labels
            self.buy_price_label.setText(f"Buy Price ({symbol}): {buy_price_formatted}")
            self.sell_price_label.setText(f"Sell Price ({symbol}): {sell_price_formatted}")
        else:
            self.add_log("No symbol selected in the combobox")

    def update_pivot_data(self):
        global last_pivot_high_value, last_pivot_low_value
        for symbol in self.symbol_settings:
            settings = self.symbol_settings[symbol]

            if settings['use_pivot']:
                url = "http://127.0.0.1:5000/receive_pivot_data"  # Local Flask server URL
                try:
                    response = requests.get(url)
                    if response.status_code == 200:
                        data = response.json()

                        new_pivot_high = data.get('last_pivot_high_value')
                        new_pivot_low = data.get('last_pivot_low_value')

                        if new_pivot_high != last_pivot_high_value:
                            last_pivot_high_value = new_pivot_high
                            settings['sell_trade_executed'] = False  
                            self.add_log(f"Pivot high for {symbol} updated to {new_pivot_high}. Reset sell_trade_executed flag.")
                        
                        if new_pivot_low != last_pivot_low_value:
                            last_pivot_low_value = new_pivot_low
                            settings['buy_trade_executed'] = False
                            self.add_log(f"Pivot low for {symbol} updated to {new_pivot_low}. Reset buy_trade_executed flag.")
                        
                    else:
                        self.add_log(f"Failed to get pivot data. Status code: {response.status_code}")
                except Exception as e:
                    self.add_log(f"Error fetching pivot data: {e}")

    def toggle_pivot_usage(self, symbol):
        """Toggle the use of pivot prices for the given symbol."""
        settings = self.symbol_settings[symbol]
        settings['use_pivot'] = not settings['use_pivot']
        status = "On" if settings['use_pivot'] else "Off"
        self.add_log(f"Pivot usage for {symbol} set to {status}")
        # Update the button appearance
        pivot_button = settings.get('pivot_toggle_button')
        if pivot_button:
            pivot_button.setText(f"Use Pivot: {status}")
            if settings['use_pivot']:
                pivot_button.setStyleSheet("background-color: green; color: white;")
            else:
                pivot_button.setStyleSheet("background-color: red; color: white;")

    def toggle_new_trade_permission(self, symbol):
        """Toggle whether new trades can be opened when an existing trade is open."""
        settings = self.symbol_settings[symbol]
        settings['allow_new_trade'] = not settings['allow_new_trade']
        status = "On" if settings['allow_new_trade'] else "Off"
        self.add_log(f"New trades permission for {symbol} set to {status}")
        # Update the button appearance
        new_trade_button = settings.get('allow_new_trade_button')
        if new_trade_button:
            new_trade_button.setText(f"Allow New Trade: {status}")
            if settings['allow_new_trade']:
                new_trade_button.setStyleSheet("background-color: green; color: white;")
            else:
                new_trade_button.setStyleSheet("background-color: red; color: white;")

    def manual_trade(self, symbol, direction, risk_entry, distance_entry, volume_entry, martingale_entry):
        try:
            balance = mt5.account_info().equity
            contract_size = mt5.symbol_info(symbol).trade_contract_size
            risk = float(risk_entry.text())
            distance_price = float(distance_entry.text())
            volume = float(volume_entry.text())
            martingale_multiplier = float(martingale_entry.text())

            if risk > 0 and distance_price > 0 and contract_size > 0:
                lot = balance * (risk / 100) / (distance_price * contract_size)
                lot = round(lot, 2)
            else:
                lot = volume  # Use static lot size if risk is zero

            adjusted_lot = round(lot, 2)
            if adjusted_lot < mt5.symbol_info(symbol).volume_min:
                adjusted_lot = mt5.symbol_info(symbol).volume_min
                self.add_log(f"Adjusted lot size is below minimum. Set to minimum lot size: {adjusted_lot}")

            self.open_position(symbol, adjusted_lot, direction)
            self.add_log(f"Manual {direction} trade executed for {symbol} with adjusted lot size {adjusted_lot}")
            self.total_profits[symbol] = 0
        except Exception as e:
            self.add_log(f"Error in manual_trade: {e}")

    def manual_reverse(self, symbol, risk_entry, distance_entry, volume_entry, martingale_entry):
        try:
            position = next((pos for pos in mt5.positions_get(symbol=symbol)), None)
            if position:
                balance = mt5.account_info().equity
                contract_size = mt5.symbol_info(symbol).trade_contract_size
                risk = float(risk_entry.text())
                distance_price = float(distance_entry.text())
                volume = float(volume_entry.text())
                martingale_multiplier = float(martingale_entry.text())

                if risk > 0 and distance_price > 0 and contract_size > 0:
                    lot = balance * (risk / 100) / (distance_price * contract_size)
                    lot = round(lot, 2)
                else:
                    lot = volume  # Use static lot size if risk is zero

                adjusted_volume = round(lot * martingale_multiplier, 2)
                if adjusted_volume < mt5.symbol_info(symbol).volume_min:
                    adjusted_volume = mt5.symbol_info(symbol).volume_min

                direction = "Sell" if position.type == mt5.ORDER_TYPE_BUY else "Buy"
                symbol_settings = self.symbol_settings.get(symbol, {'commission': self.commission})
                commission = symbol_settings['commission']
                if symbol in self.total_profits:
                    self.total_profits[symbol] += position.profit - (position.volume * commission)
                else:
                    self.total_profits[symbol] = position.profit - (position.volume * commission)

                self.close_position(position)
                self.open_position(symbol, adjusted_volume, direction)
                self.add_log(f"Manual reverse trade executed for {symbol} with adjusted volume {adjusted_volume} and updated total profit {self.total_profits[symbol]}")
            else:
                self.add_log(f"No open position found for {symbol}")
        except Exception as e:
            self.add_log(f"Error in manual_reverse: {e}")

    def add_log(self, message):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_display.appendPlainText(f"{timestamp}: {message}")

    def closeEvent(self, event):
        # Step 1: Remove all pending orders
        try:
            orders = mt5.orders_get()
            if orders:
                for order in orders:
                    result = mt5.order_send(
                        action=mt5.TRADE_ACTION_REMOVE,
                        order=order.ticket,
                        symbol=order.symbol,
                        type=order.type
                    )
                    if result.retcode == mt5.TRADE_RETCODE_DONE:
                        self.add_log(f"Successfully removed pending order {order.symbol}, ticket {order.ticket}")
                    else:
                        self.add_log(f"Failed to remove pending order {order.symbol}, ticket {order.ticket}: {result.comment}")
            else:
                self.add_log("No pending orders found.")
        except Exception as e:
            self.add_log(f"Error while removing pending orders: {e}")

        # Step 2: Write log to file
        log_dir = "log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        log_filename = os.path.join(log_dir, datetime.now().strftime("log_%Y%m%d_%H%M%S.txt"))
        with open(log_filename, "w") as log_file:
            log_file.write(self.log_display.toPlainText())
        
        self.add_log(f"Logs saved to {log_filename}")
        
        # Step 3: Accept the event to close the platform
        event.accept()

def main():
    app = QtWidgets.QApplication([])
    dashboard = TradingDashboard()

    dashboard.read_news_from_excel()

    dashboard.show()
    dashboard.update_gui_loop()
    app.exec()

if __name__ == "__main__":
    main()

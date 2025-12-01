import pandas as pd
import pyodbc
from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QPushButton, QFileDialog, QMessageBox, QTabWidget, QHBoxLayout, QInputDialog, QLabel, QDateEdit, QLineEdit, QComboBox, QTextEdit, QStyledItemDelegate, QGroupBox, QFrame
import sys
from datetime import datetime, timedelta
import numpy as np
from PyQt5.QtGui import QColor, QFont
import warnings
import os
from PyQt5.QtCore import QDate, Qt, QObject, QEvent, QObject, QEvent

#===============自訂模組======================
from bargaining import process_bargain_records, generate_settlement_voucher, generate_trading_slip, calculate_new_trade_batch, bargain_sell, generate_bargain_upload_file
from db_access import get_contracts_from_sell_table, get_631_Monitor_Fill, get_customer_bank_and_email, get_trust_info
from format_utils import strip_trailing_zeros, float_to_str_maxlen, next_business_day, strip_whitespace, edate, cusid_to_padded, format_number_to_11
from file_reader import get_daily_bond_rate, load_quote, save_trading_statement, read_today_trade_buy, read_today_trade_sell
from execution import (setup_exercise_input_search,
                      query_exercise_info, fetch_exercise_contracts,
                      update_exercise_result_table, add_exercise_info, add_exercise_to_sell)
from expired import query_expired_contracts, add_expired_to_sell
from quote_calculator import QuoteCalculatorWindow
from quote_table import QuoteTableWindow
from option_renewal import (query_renewal_contracts, add_renewal_contract, 
                           update_renewal_table, transfer_renewal_data)
from file_generator import generate_trade_notice_template
from envs import trade_notice_dir
from backend_process import (send_today_trade_email, send_bargain_trade_email, 
                             generate_today_detail, generate_trade_confirmation,
                             send_control_table_email, send_customer_positions_email,
                             clear_output_window, send_customer_detail_email)
#=============================================

warnings.filterwarnings('ignore')  # 忽略所有警告
pd.options.mode.chained_assignment = None  # 忽略 pandas 的 SettingWithCopyWarning

# 設定 pandas 顯示選項，避免資料被截斷
pd.set_option('display.max_rows', None)      # 顯示所有行
pd.set_option('display.max_columns', None)   # 顯示所有列
pd.set_option('display.width', None)         # 不限制顯示寬度
pd.set_option('display.max_colwidth', None)  # 不限制列寬度
pd.set_option('display.expand_frame_repr', False)  # 不換行顯示

print("讀取資料啟動中")

# 讀取日指標公債利率（全域變數，供所有函數使用）
rf = get_daily_bond_rate()
tday = datetime.now()

class CustomerIDComboBoxDelegate(QStyledItemDelegate):
    """客戶ID下拉式選單委託類 - 簡化版"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
    
    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        editor.setEditable(True)
        editor.setInsertPolicy(QComboBox.NoInsert)
        
        # 獲取客戶列表
        customer_list = self.parent.get_customer_list()
        for customer in customer_list:
            editor.addItem(customer)
        
        return editor
    
    def setEditorData(self, editor, index):
        value = index.data()
        if value:
            editor.setCurrentText(str(value))
    
    def setModelData(self, editor, model, index):
        # 提取客戶ID（取分隔符號前的部分）
        text = editor.currentText()
        if " - " in text:
            customer_id = text.split(" - ")[0]
        else:
            customer_id = text
        model.setData(index, customer_id)

class BuySellComboBoxDelegate(QStyledItemDelegate):
    """買/賣下拉式選單委託類"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
    
    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        editor.setEditable(False)  # 不可編輯，只能選擇
        editor.setInsertPolicy(QComboBox.NoInsert)
        
        # 添加買/賣選項
        editor.addItem("買")
        editor.addItem("賣")
        
        return editor
    
    def setEditorData(self, editor, index):
        value = index.data()
        if value:
            editor.setCurrentText(str(value))
    
    def setModelData(self, editor, model, index):
        text = editor.currentText()
        model.setData(index, text)

class TableWidgetWithDelete(QTableWidget):
    """支援鍵盤刪除的表格類"""
    def keyPressEvent(self, event):
        # 檢查是否按下 Delete 或 Backspace 鍵
        if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            # 獲取選中的行
            selected_rows = set()
            selected_items = self.selectedItems()
            if selected_items:
                for item in selected_items:
                    if item:
                        selected_rows.add(item.row())
                
                if selected_rows:
                    # 如果表格啟用了排序，需要先禁用排序再刪除
                    was_sorting_enabled = self.isSortingEnabled()
                    if was_sorting_enabled:
                        self.setSortingEnabled(False)
                    
                    # 按照從高到低的順序刪除行（避免索引變化問題）
                    for row in sorted(selected_rows, reverse=True):
                        if 0 <= row < self.rowCount():
                            self.removeRow(row)
                    
                    # 如果之前啟用了排序，重新啟用
                    if was_sorting_enabled:
                        self.setSortingEnabled(True)
                    
                    return  # 事件已處理
        # 調用父類的方法處理其他按鍵
        super().keyPressEvent(event)

class ReferencePriceComboBoxDelegate(QStyledItemDelegate):
    """參考價下拉式選單委託類"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
    
    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        editor.setEditable(False)
        editor.setInsertPolicy(QComboBox.NoInsert)
        editor.addItems(["開盤參考價", "當日收盤價", "盤中市場價"])
        return editor
    
    def setEditorData(self, editor, index):
        value = index.data()
        if value:
            editor.setCurrentText(str(value))
    
    def setModelData(self, editor, model, index):
        text = editor.currentText()
        model.setData(index, text)

class RecordingPersonComboBoxDelegate(QStyledItemDelegate):
    """錄音人員下拉式選單委託類"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
    
    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        editor.setEditable(False)
        editor.setInsertPolicy(QComboBox.NoInsert)
        
        # 添加錄音人員選項，默認值為"蔡睿"
        editor.addItem("蔡睿")
        editor.addItem("詹郁文")
        editor.addItem("王慕約")
        editor.addItem("黃暐庭")
        editor.addItem("朱軒慧")
        
        return editor
    
    def setEditorData(self, editor, index):
        value = index.data()
        if value:
            # 如果已有值，設置為該值
            editor.setCurrentText(str(value))
        else:
            # 如果沒有值，設置默認值為"蔡睿"
            editor.setCurrentText("蔡睿")
    
    def setModelData(self, editor, model, index):
        text = editor.currentText()
        model.setData(index, text)

#====================類外函數已移至類內部===================
class TableEditor(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("CBAS 交易資料編輯器")
        self.resize(2000, 800)
        self.tabs = QTabWidget()
        # ===== 新增交割日 QDateEdit =====
        self.label_settle = QLabel("交割日：")
        self.dateedit_settle = QDateEdit()
        self.dateedit_settle.setCalendarPopup(True)
        # 預設值為 T+2 business day
        tday = datetime.now()
        default_settle = next_business_day(tday, 2)
        self.dateedit_settle.setDate(QDate(default_settle.year, default_settle.month, default_settle.day))
        # 集中定義分頁名稱
        self.tab_names = ["交易處理", "後台工作"]

        # 交易處理分頁
        trading_widget = QWidget()
        trading_layout = QVBoxLayout()
        
        # 建立交易處理的子分頁
        trading_tabs = QTabWidget()
        
        # 新作買進子分頁
        buy_widget = QWidget()
        buy_layout = QVBoxLayout()
        
        # 新作買進按鈕區域
        buy_btn_frame = QWidget()
        buy_btn_layout = QHBoxLayout()
        
        # 新作買進相關按鈕
        self.btn_buy_add = QPushButton("新增")
        self.btn_buy_add.setFixedSize(80, 30)
        self.btn_buy_add.clicked.connect(lambda: self.add_row_specific("新作買進"))
        
        self.btn_buy_delete = QPushButton("刪除")
        self.btn_buy_delete.setFixedSize(80, 30)
        self.btn_buy_delete.clicked.connect(lambda: self.delete_row_specific("新作買進"))
        # 新增：報價確認 按鈕
        self.btn_buy_quote_check = QPushButton("報價確認")
        self.btn_buy_quote_check.setFixedSize(80, 30)
        self.btn_buy_quote_check.clicked.connect(self.check_buy_table_with_quote)
        
        self.btn_buy_qty_check = QPushButton("張數/總額確認")
        self.btn_buy_qty_check.setFixedSize(120, 30)
        self.btn_buy_qty_check.clicked.connect(self.check_buy_table_with_qty)
        
        self.btn_buy_generate = QPushButton("產生上傳檔")
        self.btn_buy_generate.setFixedSize(100, 30)
        self.btn_buy_generate.clicked.connect(self.generate_buy_upload_file)
        

        
        # 添加到按鈕布局
        buy_btn_layout.addWidget(self.btn_buy_add)
        buy_btn_layout.addWidget(self.btn_buy_delete)
        buy_btn_layout.addWidget(self.btn_buy_quote_check)
        buy_btn_layout.addWidget(self.btn_buy_qty_check)
        buy_btn_layout.addWidget(self.btn_buy_generate)
        
        # 新增：重新編號按鈕
        self.btn_buy_renumber = QPushButton("重新編號")
        self.btn_buy_renumber.setFixedSize(100, 30)
        self.btn_buy_renumber.clicked.connect(self.renumber_buy_table)

        buy_btn_layout.addWidget(self.btn_buy_renumber)
        buy_btn_layout.addStretch()
        
        buy_btn_frame.setLayout(buy_btn_layout)
        
        # 新作買進表格
        self.table_buy = TableWidgetWithDelete()
        buy_columns = ['新作契約編號', '上傳序號', '交易類型', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '成交張數', '履約利率%', '成交均價', 
            '權利金百元價', '手續費(業務單位)', '成交金額', '單位權利金', '權利金總額', '錄音日期', '錄音時間', '交易日', '生效日', '交割日期', '選擇權到期日', '賣回日', '賣回價', '提前履約界限日', '提前履約賠償金', '標的波動率', '無風險利率',
            '資金成本', '轉債面額', '選擇權型態', '選擇權買賣別', '報價方式', '短契約', '手續費(營業員)', '交易員', '營業員', '錄音人員', '子帳號', '固定端契約編號', '長約附加條款', '價格事件', '來自']
        self.table_buy.setColumnCount(len(buy_columns))
        self.table_buy.setHorizontalHeaderLabels(buy_columns)
        self.table_buy.setSortingEnabled(True)  # 啟用欄位排序
        
        # 添加到主布局
        buy_layout.addWidget(buy_btn_frame)
        buy_layout.addWidget(self.table_buy)
        buy_widget.setLayout(buy_layout)

        # 提解賣出分頁
        sell_widget = QWidget()
        sell_layout = QVBoxLayout()
        
        # 提解賣出按鈕區域
        sell_btn_frame = QWidget()
        sell_btn_layout = QHBoxLayout()
        
        # 提解賣出相關按鈕
        self.btn_sell_add = QPushButton("新增")
        self.btn_sell_add.setFixedSize(80, 30)
        self.btn_sell_add.clicked.connect(lambda: self.add_row_specific("提解賣出"))
        
        self.btn_sell_delete = QPushButton("刪除")
        self.btn_sell_delete.setFixedSize(80, 30)
        self.btn_sell_delete.clicked.connect(lambda: self.delete_row_specific("提解賣出"))
        
        # 新增：張數/總額確認 按鈕
        self.btn_sell_qty_check = QPushButton("張數/總額確認")
        self.btn_sell_qty_check.setFixedSize(120, 30)
        self.btn_sell_qty_check.clicked.connect(self.check_sell_table_with_qty)
        
        # 新增：產已實CBAS契約 按鈕
        self.btn_sell_realized = QPushButton("產已實CBAS契約")
        self.btn_sell_realized.setFixedSize(120, 30)
        self.btn_sell_realized.clicked.connect(self.generate_i_realized_file)

        self.btn_sell_generate = QPushButton("產生上傳檔")
        self.btn_sell_generate.setFixedSize(100, 30)
        self.btn_sell_generate.clicked.connect(self.generate_sell_upload_file)
        
        # 添加到按鈕布局
        sell_btn_layout.addWidget(self.btn_sell_add)
        sell_btn_layout.addWidget(self.btn_sell_delete)
        sell_btn_layout.addWidget(self.btn_sell_qty_check)
        sell_btn_layout.addWidget(self.btn_sell_realized)
        sell_btn_layout.addWidget(self.btn_sell_generate)
        
        # 新增：重新編號按鈕
        self.btn_sell_renumber = QPushButton("重新編號")
        self.btn_sell_renumber.setFixedSize(100, 30)
        self.btn_sell_renumber.clicked.connect(self.renumber_sell_table)

        sell_btn_layout.addWidget(self.btn_sell_renumber)
        sell_btn_layout.addStretch()
        
        sell_btn_frame.setLayout(sell_btn_layout)
        
        # 提解賣出表格
        self.table_sell = TableWidgetWithDelete()
        sell_columns = [
            '解約契約編號', '原單契約編號', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
            '履約張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日', '提前履約賠償金', '履約價', 
            '選擇權交割單價', '交割總金額', '錄音時間', '來自'
        ]
        self.table_sell.setColumnCount(len(sell_columns))
        self.table_sell.setHorizontalHeaderLabels(sell_columns)
        self.table_sell.setSortingEnabled(True)  # 啟用欄位排序
        
        # 添加到主布局
        sell_layout.addWidget(sell_btn_frame)
        sell_layout.addWidget(self.table_sell)
        sell_widget.setLayout(sell_layout)

        # 錄音分頁
        recording_widget = QWidget()
        recording_layout = QVBoxLayout()
        
        # 錄音按鈕區域
        recording_btn_frame = QWidget()
        recording_btn_layout = QHBoxLayout()
        
        # 錄音相關按鈕
        self.btn_recording_refresh = QPushButton("刷新資料")
        self.btn_recording_refresh.setFixedSize(100, 30)
        self.btn_recording_refresh.clicked.connect(self.refresh_recording_table)
        
        self.btn_recording_generate = QPushButton("產生錄音檔")
        self.btn_recording_generate.setFixedSize(100, 30)
        self.btn_recording_generate.clicked.connect(self.generate_recording_file)
        
        # 添加到按鈕布局
        recording_btn_layout.addWidget(self.btn_recording_refresh)
        recording_btn_layout.addWidget(self.btn_recording_generate)
        recording_btn_layout.addStretch()
        
        recording_btn_frame.setLayout(recording_btn_layout)
        
        # 錄音表格
        self.table_recording = TableWidgetWithDelete()
        recording_columns = ['客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '買進張數', '賣出張數', '成交均價', 'CELLPHONE', '授權人', '授權人電話', '錄音時間', '錄音人員']
        self.table_recording.setColumnCount(len(recording_columns))
        self.table_recording.setHorizontalHeaderLabels(recording_columns)
        self.table_recording.setSortingEnabled(True)  # 啟用欄位排序
        
        # 設定錄音人員欄位的編輯器（第12欄，索引11）
        recording_person_col_idx = recording_columns.index('錄音人員')
        self.table_recording.setItemDelegateForColumn(recording_person_col_idx, RecordingPersonComboBoxDelegate(self))
        
        # 添加到主布局
        recording_layout.addWidget(recording_btn_frame)
        recording_layout.addWidget(self.table_recording)
        recording_widget.setLayout(recording_layout)

        # VIP名單分頁
        vip_list_widget = QWidget()
        vip_list_layout = QVBoxLayout()
        
        # VIP名單按鈕區域
        vip_list_btn_frame = QWidget()
        vip_list_btn_layout = QHBoxLayout()
        
        # VIP名單相關按鈕
        self.btn_vip_list_add = QPushButton("新增")
        self.btn_vip_list_add.setFixedSize(80, 30)
        self.btn_vip_list_add.clicked.connect(lambda: self.add_row_specific("VIP名單"))
        
        self.btn_vip_list_delete = QPushButton("刪除")
        self.btn_vip_list_delete.setFixedSize(80, 30)
        self.btn_vip_list_delete.clicked.connect(lambda: self.delete_row_specific("VIP名單"))
        
        self.btn_vip_list_save = QPushButton("儲存")
        self.btn_vip_list_save.setFixedSize(80, 30)
        self.btn_vip_list_save.clicked.connect(self.save_vip_list)
        
        # 添加到按鈕布局
        vip_list_btn_layout.addWidget(self.btn_vip_list_add)
        vip_list_btn_layout.addWidget(self.btn_vip_list_delete)
        vip_list_btn_layout.addWidget(self.btn_vip_list_save)
        vip_list_btn_layout.addStretch()
        
        vip_list_btn_frame.setLayout(vip_list_btn_layout)
        
        # VIP名單表格
        self.table_vip_list = TableWidgetWithDelete()
        vip_list_columns = ['客戶ID', '客戶名稱', '不限張數低手續費', '不限張數低利率']
        self.table_vip_list.setColumnCount(len(vip_list_columns))
        self.table_vip_list.setHorizontalHeaderLabels(vip_list_columns)
        
        # 添加到主布局
        vip_list_layout.addWidget(vip_list_btn_frame)
        vip_list_layout.addWidget(self.table_vip_list)
        vip_list_widget.setLayout(vip_list_layout)

        # 特殊報價分頁
        vip_quote_widget = QWidget()
        vip_quote_layout = QVBoxLayout()
        
        # 特殊報價按鈕區域
        vip_quote_btn_frame = QWidget()
        vip_quote_btn_layout = QHBoxLayout()
        
        # 特殊報價相關按鈕
        self.btn_vip_quote_add = QPushButton("新增")
        self.btn_vip_quote_add.setFixedSize(80, 30)
        self.btn_vip_quote_add.clicked.connect(lambda: self.add_row_specific("特殊報價"))
        
        self.btn_vip_quote_delete = QPushButton("刪除")
        self.btn_vip_quote_delete.setFixedSize(80, 30)
        self.btn_vip_quote_delete.clicked.connect(lambda: self.delete_row_specific("特殊報價"))
        
        self.btn_vip_quote_save = QPushButton("儲存")
        self.btn_vip_quote_save.setFixedSize(80, 30)
        self.btn_vip_quote_save.clicked.connect(self.save_vip_quote)
        
        # 添加到按鈕布局
        vip_quote_btn_layout.addWidget(self.btn_vip_quote_add)
        vip_quote_btn_layout.addWidget(self.btn_vip_quote_delete)
        vip_quote_btn_layout.addWidget(self.btn_vip_quote_save)
        vip_quote_btn_layout.addStretch()
        
        vip_quote_btn_frame.setLayout(vip_quote_btn_layout)
        
        # 特殊報價表格
        self.table_vip_quote = TableWidgetWithDelete()
        vip_quote_columns = ['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '利率%', '手續費']
        self.table_vip_quote.setColumnCount(len(vip_quote_columns))
        self.table_vip_quote.setHorizontalHeaderLabels(vip_quote_columns)
        
        # 添加到主布局
        vip_quote_layout.addWidget(vip_quote_btn_frame)
        vip_quote_layout.addWidget(self.table_vip_quote)
        vip_quote_widget.setLayout(vip_quote_layout)

        # 議價交易分頁
        bargain_widget = QWidget()
        bargain_layout = QVBoxLayout()
        
        # 議價交易按鈕區域
        bargain_btn_frame = QWidget()
        bargain_btn_layout = QHBoxLayout()
        
        # 議價交易相關按鈕
        self.btn_process_bargain = QPushButton("處理議價")
        self.btn_process_bargain.setFixedSize(120, 40)
        self.btn_process_bargain.clicked.connect(self.process_bargain)
        
        self.btn_generate_tickets = QPushButton("產給付憑證及買賣成交單")
        self.btn_generate_tickets.setFixedSize(180, 40)
        self.btn_generate_tickets.clicked.connect(self.generate_tickets)
        
        self.btn_add_to_new_trade = QPushButton("新增至新作/提解")
        self.btn_add_to_new_trade.setFixedSize(120, 40)
        self.btn_add_to_new_trade.clicked.connect(self.add_bargain_to_new_trade)
        
        # 新增按鈕
        self.btn_add_bargain = QPushButton("新增")
        self.btn_add_bargain.setFixedSize(80, 40)
        self.btn_add_bargain.clicked.connect(lambda: self.add_row_specific("議價交易"))
        
        # 刪除按鈕
        self.btn_delete_bargain = QPushButton("刪除")
        self.btn_delete_bargain.setFixedSize(80, 40)
        self.btn_delete_bargain.clicked.connect(lambda: self.delete_row_specific("議價交易"))
        
        # 添加到按鈕布局
        bargain_btn_layout.addWidget(self.btn_add_bargain)
        bargain_btn_layout.addWidget(self.btn_delete_bargain)
        bargain_btn_layout.addWidget(self.btn_process_bargain)
        bargain_btn_layout.addWidget(self.btn_generate_tickets)
        bargain_btn_layout.addWidget(self.btn_add_to_new_trade)
        bargain_btn_layout.addStretch()
        
        bargain_btn_frame.setLayout(bargain_btn_layout)
        
        # 議價交易表格
        self.table_bargain = TableWidgetWithDelete()
        bargain_columns = ['單據編號', '成交日期', 'T+?交割', '買/賣','客戶ID', 'CB代號', '議價張數', '議價價格','參考價', '錄音時間', '交割日期', '客戶名稱', 'CB名稱', '議價金額', '備註', '銀行', '分行', '銀行帳號', '集保帳號', '通訊地址']
        self.table_bargain.setColumnCount(len(bargain_columns))
        self.table_bargain.setHorizontalHeaderLabels(bargain_columns)
        
        # 設定買/賣欄位的編輯器（第4欄，索引3）
        self.table_bargain.setItemDelegateForColumn(3, BuySellComboBoxDelegate(self))
        # 設定客戶ID欄位的編輯器（第5欄，索引4）
        self.table_bargain.setItemDelegateForColumn(4, CustomerIDComboBoxDelegate(self))
        # 設定參考價欄位的編輯器（第9欄，索引8）
        self.table_bargain.setItemDelegateForColumn(8, ReferencePriceComboBoxDelegate(self))
        
        # 添加到主布局
        bargain_layout.addWidget(QLabel("議價交易處理"))
        bargain_layout.addWidget(bargain_btn_frame)
        bargain_layout.addWidget(QLabel("議價交易資料"))
        bargain_layout.addWidget(self.table_bargain)
        
        bargain_widget.setLayout(bargain_layout)

        # 實物履約分頁
        cbas_to_cb_widget = QWidget()
        cbas_to_cb_layout = QVBoxLayout()
        
        # 輸入區域
        input_frame = QWidget()
        input_layout = QHBoxLayout()
        
        # 四個輸入欄位
        self.label_cus_id = QLabel("客戶ID：")
        self.input_cus_id = QComboBox()
        self.input_cus_id.setEditable(True)
        self.input_cus_id.setInsertPolicy(QComboBox.NoInsert)
        self.input_cus_id.setPlaceholderText("請輸入客戶ID")
        
        self.label_cb_code = QLabel("CB代號：")
        self.input_cb_code = QComboBox()
        self.input_cb_code.setEditable(True)
        self.input_cb_code.setInsertPolicy(QComboBox.NoInsert)
        self.input_cb_code.setPlaceholderText("請輸入CB代號")
        
        self.label_exercise_qty = QLabel("履約張數：")
        self.input_exercise_qty = QLineEdit()
        self.input_exercise_qty.setPlaceholderText("請輸入履約張數")
        
        self.label_settlement_date = QLabel("交割日：")
        self.input_settlement_date = QDateEdit()
        self.input_settlement_date.setCalendarPopup(True)
        # 預設值為 T+1 business day
        today = datetime.now()
        default_settlement = next_business_day(today, 1)
        self.input_settlement_date.setDate(QDate(default_settlement.year, default_settlement.month, default_settlement.day))
        
        # 查詢按鈕
        self.btn_query_exercise = QPushButton("查詢履約資訊")
        self.btn_query_exercise.setFixedSize(120, 30)
        self.btn_query_exercise.clicked.connect(lambda: query_exercise_info(
            self.input_cus_id, self.input_cb_code, self.input_exercise_qty, 
            self.input_settlement_date, self.df_quote, self.table_exercise_result, 
            fetch_exercise_contracts
        ))
        
        # 新增履約資訊按鈕
        self.btn_add_exercise = QPushButton("新增履約資訊")
        self.btn_add_exercise.setFixedSize(120, 30)
        self.btn_add_exercise.clicked.connect(lambda: add_exercise_info(
            self.table_exercise_result, self.table_cbas_to_cb, self.get_table_data
        ))
        
        # 新增至提解按鈕
        self.btn_add_to_sell = QPushButton("新增至提解")
        self.btn_add_to_sell.setFixedSize(120, 30)
        self.btn_add_to_sell.clicked.connect(lambda: add_exercise_to_sell(
            self.table_cbas_to_cb, self.df_quote, self.get_table_data, self.show_sell_table
        ))
        
        # 刪除按鈕
        self.btn_delete_exercise = QPushButton("刪除")
        self.btn_delete_exercise.setFixedSize(80, 30)
        self.btn_delete_exercise.clicked.connect(lambda: self.delete_row_specific("實物履約"))
        
        # 設置實物履約輸入框的搜尋功能（將在UI創建完成後調用）
        
        # 添加到輸入布局
        input_layout.addWidget(self.label_cus_id)
        input_layout.addWidget(self.input_cus_id)
        input_layout.addWidget(self.label_cb_code)
        input_layout.addWidget(self.input_cb_code)
        input_layout.addWidget(self.label_exercise_qty)
        input_layout.addWidget(self.input_exercise_qty)
        input_layout.addWidget(self.label_settlement_date)
        input_layout.addWidget(self.input_settlement_date)
        input_layout.addWidget(self.btn_query_exercise)
        input_layout.addWidget(self.btn_add_exercise)
        input_layout.addWidget(self.btn_add_to_sell)
        input_layout.addWidget(self.btn_delete_exercise)
        input_layout.addStretch()
        
        input_frame.setLayout(input_layout)
        
        # 結果顯示表格
        self.table_exercise_result = TableWidgetWithDelete()
        exercise_result_columns = ['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '原單契約編號', '原利率', '交易日期', '交割日期', '成交日期', '履約價', '賣出金額', '原庫存張數', '今日賣出張數', '履約張數', '履約後剩餘張數', '解約類別', '履約方式']
        self.table_exercise_result.setColumnCount(len(exercise_result_columns))
        self.table_exercise_result.setHorizontalHeaderLabels(exercise_result_columns)
        
        # 本日新增履約表格
        self.table_cbas_to_cb = TableWidgetWithDelete()
        cbas_to_cb_columns = ['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '原單契約編號', '原利率', '交易日期', '交割日期', '成交日期', '履約價', '賣出金額', '原庫存張數', '今日賣出張數', '履約張數', '履約後剩餘張數', '解約類別', '履約方式', '備註', '錄音時間']
        self.table_cbas_to_cb.setColumnCount(len(cbas_to_cb_columns))
        self.table_cbas_to_cb.setHorizontalHeaderLabels(cbas_to_cb_columns)
        
        # 添加到主布局
        cbas_to_cb_layout.addWidget(QLabel("實物履約查詢"))
        cbas_to_cb_layout.addWidget(input_frame)
        cbas_to_cb_layout.addWidget(QLabel("查詢結果"))
        cbas_to_cb_layout.addWidget(self.table_exercise_result)
        cbas_to_cb_layout.addWidget(QLabel("本日新增履約"))
        cbas_to_cb_layout.addWidget(self.table_cbas_to_cb)
        
        cbas_to_cb_widget.setLayout(cbas_to_cb_layout)

        # 合約到期分頁
        expired_widget = QWidget()
        expired_layout = QVBoxLayout()
        
        # 日期選擇區域
        date_frame = QWidget()
        date_layout = QHBoxLayout()
        
        # 到期日選擇器
        date_layout.addWidget(QLabel("到期日："))
        self.dateedit_expired = QDateEdit()
        self.dateedit_expired.setDate(QDate.currentDate())  # 預設為今日
        self.dateedit_expired.setCalendarPopup(True)  # 啟用日曆彈出
        self.dateedit_expired.setFixedSize(120, 30)
        date_layout.addWidget(self.dateedit_expired)
        date_layout.addStretch()
        
        date_frame.setLayout(date_layout)
        
        # 查詢按鈕區域
        expired_btn_frame = QWidget()
        expired_btn_layout = QHBoxLayout()
        
        # 查詢按鈕（修改名稱）
        self.btn_query_expired = QPushButton("查詢")
        self.btn_query_expired.setFixedSize(120, 40)
        self.btn_query_expired.clicked.connect(lambda: query_expired_contracts(self.dateedit_expired, self.table_expired, self.df_quote, self.table_sell, self.get_table_data))
        
        # 新增至提解按鈕
        self.btn_add_expired_to_sell = QPushButton("新增至提解")
        self.btn_add_expired_to_sell.setFixedSize(120, 40)
        self.btn_add_expired_to_sell.clicked.connect(lambda: add_expired_to_sell(self.table_expired, self.get_table_data, self.show_sell_table))
        
        # 刪除按鈕
        self.btn_delete_expired = QPushButton("刪除")
        self.btn_delete_expired.setFixedSize(80, 40)
        self.btn_delete_expired.clicked.connect(lambda: self.delete_row_specific("合約到期"))
        
        # 添加到按鈕布局
        expired_btn_layout.addWidget(self.btn_query_expired)
        expired_btn_layout.addWidget(self.btn_add_expired_to_sell)
        expired_btn_layout.addWidget(self.btn_delete_expired)
        expired_btn_layout.addStretch()
        
        expired_btn_frame.setLayout(expired_btn_layout)
        
        # 合約到期表格
        self.table_expired = TableWidgetWithDelete()
        expired_columns = [
            '原單契約編號', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
            '原庫存張數', '今日賣出張數', '剩餘到期張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日',
            '提前履約賠償金', '履約價', '選擇權交割單價', '交割總金額', '錄音時間'
        ]
        self.table_expired.setColumnCount(len(expired_columns))
        self.table_expired.setHorizontalHeaderLabels(expired_columns)
        
        # 添加到主布局
        expired_layout.addWidget(QLabel("合約到期查詢"))
        expired_layout.addWidget(date_frame)
        expired_layout.addWidget(expired_btn_frame)
        expired_layout.addWidget(QLabel("到期合約資料"))
        expired_layout.addWidget(self.table_expired)
        
        expired_widget.setLayout(expired_layout)

        # 選擇權續期分頁
        renewal_widget = QWidget()
        renewal_layout = QVBoxLayout()
        
        # 查詢輸入區域
        renewal_input_frame = QWidget()
        renewal_input_layout = QHBoxLayout()
        
        # 客戶ID輸入框
        self.label_renewal_cus_id = QLabel("客戶ID：")
        self.input_renewal_cus_id = QComboBox()
        self.input_renewal_cus_id.setEditable(True)
        self.input_renewal_cus_id.setInsertPolicy(QComboBox.NoInsert)
        self.input_renewal_cus_id.setPlaceholderText("請輸入客戶ID")
        
        # CB代號輸入框
        self.label_renewal_cb_code = QLabel("CB代號：")
        self.input_renewal_cb_code = QComboBox()
        self.input_renewal_cb_code.setEditable(True)
        self.input_renewal_cb_code.setInsertPolicy(QComboBox.NoInsert)
        self.input_renewal_cb_code.setPlaceholderText("請輸入CB代號")
        
        # 續期查詢按鈕
        self.btn_query_renewal = QPushButton("續期查詢")
        self.btn_query_renewal.setFixedSize(120, 30)
        self.btn_query_renewal.clicked.connect(lambda: self.query_renewal_contracts())
        
        # 新增續期按鈕
        self.btn_add_renewal = QPushButton("新增續期")
        self.btn_add_renewal.setFixedSize(120, 30)
        self.btn_add_renewal.clicked.connect(lambda: add_renewal_contract(self.table_renewal_query, self.table_renewal_buy, self.table_renewal_sell, self.df_original_contracts, self.df_quote))
        
        # 刪除按鈕
        self.btn_delete_renewal = QPushButton("刪除")
        self.btn_delete_renewal.setFixedSize(80, 30)
        self.btn_delete_renewal.clicked.connect(lambda: self.delete_row_specific("選擇權續期"))
        
        # 設置選擇權續期輸入框的搜尋功能（將在UI創建完成後調用）
        
        # 添加到輸入布局
        renewal_input_layout.addWidget(self.label_renewal_cus_id)
        renewal_input_layout.addWidget(self.input_renewal_cus_id)
        renewal_input_layout.addWidget(self.label_renewal_cb_code)
        renewal_input_layout.addWidget(self.input_renewal_cb_code)
        renewal_input_layout.addWidget(self.btn_query_renewal)
        renewal_input_layout.addWidget(self.btn_add_renewal)
        renewal_input_layout.addWidget(self.btn_delete_renewal)
        renewal_input_layout.addStretch()
        
        renewal_input_frame.setLayout(renewal_input_layout)
        
        # 上方查詢結果表格
        self.table_renewal_query = TableWidgetWithDelete()
        renewal_query_columns = ['選擇', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '原庫存張數', '今賣出張數', '今剩餘張數', '續期張數', '今履約利率', '今成交均價']
        self.table_renewal_query.setColumnCount(len(renewal_query_columns))
        self.table_renewal_query.setHorizontalHeaderLabels(renewal_query_columns)
        # 設定選擇欄位寬度較小
        self.table_renewal_query.setColumnWidth(0, 50)  # 將第一欄 '選擇' 的寬度設為50像素
        
        # 下方左右分割區域
        bottom_frame = QWidget()
        bottom_layout = QHBoxLayout()
        
        # 左側：新作表格
        left_frame = QWidget()
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("新作"))
        
        self.table_renewal_buy = TableWidgetWithDelete()
        renewal_new_columns = ['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '續期張數', '今履約利率', '今成交均價']
        self.table_renewal_buy.setColumnCount(len(renewal_new_columns))
        self.table_renewal_buy.setHorizontalHeaderLabels(renewal_new_columns)
        left_layout.addWidget(self.table_renewal_buy)
        left_frame.setLayout(left_layout)
        
        # 右側：賣出表格
        right_frame = QWidget()
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("賣出"))
        
        self.table_renewal_sell = TableWidgetWithDelete()
        renewal_sell_columns = ['新作契約編號', '客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '原庫存張數', '今賣出張數', '續期張數', '成交均價']
        self.table_renewal_sell.setColumnCount(len(renewal_sell_columns))
        self.table_renewal_sell.setHorizontalHeaderLabels(renewal_sell_columns)
        right_layout.addWidget(self.table_renewal_sell)
        right_frame.setLayout(right_layout)
        
        # 中間：轉換按鈕
        middle_frame = QWidget()
        middle_layout = QVBoxLayout()
        # 添加一些空間讓按鈕垂直居中
        middle_layout.addStretch(1)
        middle_layout.addWidget(QLabel(""))  # 空白標籤保持對齊
        
        self.btn_transfer_renewal = QPushButton("轉換")
        self.btn_transfer_renewal.setFixedSize(80, 40)  # 設定按鈕固定大小
        self.btn_transfer_renewal.clicked.connect(lambda: transfer_renewal_data(self.table_renewal_buy, self.table_renewal_sell, self.df_original_contracts, calculate_new_trade_batch, self.show_buy_table, self.show_sell_table))
        middle_layout.addWidget(self.btn_transfer_renewal)
        
        middle_layout.addStretch(1)  # 底部也添加空間
        middle_frame.setLayout(middle_layout)
        middle_frame.setFixedWidth(100)  # 設定中間區域固定寬度
        
        # 添加左中右框到底部布局
        bottom_layout.addWidget(left_frame)
        bottom_layout.addWidget(middle_frame)
        bottom_layout.addWidget(right_frame)
        bottom_frame.setLayout(bottom_layout)
        
        # 添加到主布局
        renewal_layout.addWidget(QLabel("選擇權續期查詢"))
        renewal_layout.addWidget(renewal_input_frame)
        renewal_layout.addWidget(QLabel("查詢結果"))
        renewal_layout.addWidget(self.table_renewal_query)
        renewal_layout.addWidget(bottom_frame)
        
        renewal_widget.setLayout(renewal_layout)

        # 常用客戶維護分頁
        customer_widget = QWidget()
        customer_layout = QVBoxLayout()
        
        # 常用客戶維護按鈕區域
        customer_btn_frame = QWidget()
        customer_btn_layout = QHBoxLayout()
        
        # 常用客戶維護相關按鈕
        self.btn_customer_add = QPushButton("新增")
        self.btn_customer_add.setFixedSize(80, 30)
        self.btn_customer_add.clicked.connect(lambda: self.add_row_specific("常用客戶維護"))
        
        self.btn_customer_delete = QPushButton("刪除")
        self.btn_customer_delete.setFixedSize(80, 30)
        self.btn_customer_delete.clicked.connect(lambda: self.delete_row_specific("常用客戶維護"))
        
        self.btn_customer_save = QPushButton("儲存")
        self.btn_customer_save.setFixedSize(80, 30)
        self.btn_customer_save.clicked.connect(self.save_customer_list)
        
        # 添加到按鈕布局
        customer_btn_layout.addWidget(self.btn_customer_add)
        customer_btn_layout.addWidget(self.btn_customer_delete)
        customer_btn_layout.addWidget(self.btn_customer_save)
        customer_btn_layout.addStretch()
        
        customer_btn_frame.setLayout(customer_btn_layout)
        
        # 常用客戶維護表格
        self.table_customer = TableWidgetWithDelete()
        customer_columns = ['客戶ID', '客戶名稱']
        self.table_customer.setColumnCount(len(customer_columns))
        self.table_customer.setHorizontalHeaderLabels(customer_columns)
        
        # 添加到主布局
        customer_layout.addWidget(customer_btn_frame)
        customer_layout.addWidget(self.table_customer)
        customer_widget.setLayout(customer_layout)

        # 用分頁名稱與 widget 對應
        self.tab_widgets = [buy_widget, sell_widget, vip_list_widget, vip_quote_widget, bargain_widget, cbas_to_cb_widget, expired_widget, renewal_widget] #名字要再改
        for name, widget in zip(self.tab_names, self.tab_widgets):
            self.tabs.addTab(widget, name)

        # 全局按鈕（只保留讀取今日買賣和重新載入報價表）
        self.btn_refresh = QPushButton("讀取今日買賣")
        self.btn_refresh_quote = QPushButton("重新載入報價表")
        self.btn_open_quote = QPushButton("打開報價表")
        self.btn_refresh.setFixedSize(120, 40)
        self.btn_refresh_quote.setFixedSize(140, 40)
        self.btn_open_quote.setFixedSize(120, 40)
        self.btn_refresh.clicked.connect(self.refresh_data)
        self.btn_refresh_quote.clicked.connect(self.refresh_quote)
        self.btn_open_quote.clicked.connect(self.open_quote_file)

        # 主 layout
        main_layout = QVBoxLayout()
        # ===== 交割日區塊 =====
        settle_layout = QHBoxLayout()
        settle_layout.addStretch()
        settle_layout.addWidget(self.label_settle)
        settle_layout.addWidget(self.dateedit_settle)
        settle_layout.addStretch()
        main_layout.addLayout(settle_layout)
        
        # 全局按鈕布局
        global_btn_layout = QHBoxLayout()
        global_btn_layout.addStretch()
        global_btn_layout.addWidget(self.btn_refresh_quote)
        global_btn_layout.addWidget(self.btn_open_quote)
        
        # 查看報價表按鈕
        self.btn_view_quote = QPushButton("查看報價表")
        self.btn_view_quote.setFixedSize(120, 40)
        self.btn_view_quote.clicked.connect(self.show_quote_table)
        global_btn_layout.addWidget(self.btn_view_quote)
        
        # 報價計算機按鈕
        self.btn_quote_calculator = QPushButton("報價計算機")
        self.btn_quote_calculator.setFixedSize(120, 40)
        self.btn_quote_calculator.clicked.connect(self.show_quote_calculator)
        global_btn_layout.addWidget(self.btn_quote_calculator)
        
        global_btn_layout.addStretch()
        global_btn_layout.setSpacing(15)
        main_layout.addLayout(global_btn_layout)
        
        # 建立交易處理分頁
        trading_widget = QWidget()
        trading_layout = QVBoxLayout()
        
        # 添加統一的暫存和讀取按鈕到左上角
        temp_btn_frame = QWidget()
        temp_btn_layout = QHBoxLayout()
        
        self.btn_temp_save = QPushButton("暫存")
        self.btn_temp_save.setFixedSize(100, 40)
        self.btn_temp_save.clicked.connect(self.temp_save_all)
        
        self.btn_temp_load = QPushButton("讀取暫存")
        self.btn_temp_load.setFixedSize(100, 40)
        self.btn_temp_load.clicked.connect(self.temp_load_all)
        
        self.btn_open_upload_folder = QPushButton("打開資料夾")
        self.btn_open_upload_folder.setFixedSize(100, 40)
        self.btn_open_upload_folder.clicked.connect(self.open_upload_folder)
        
        # 將「讀取今日買賣」放到最左邊
        temp_btn_layout.addWidget(self.btn_refresh)
        temp_btn_layout.addWidget(self.btn_temp_save)
        temp_btn_layout.addWidget(self.btn_temp_load)
        temp_btn_layout.addWidget(self.btn_open_upload_folder)
        temp_btn_layout.addStretch()
        
        temp_btn_frame.setLayout(temp_btn_layout)
        trading_layout.addWidget(temp_btn_frame)
        
        # 建立交易處理的子分頁
        trading_tabs = QTabWidget()
        
        # 將所有現有的widget添加到交易處理的子分頁中
        trading_tabs.addTab(buy_widget, "新作買進")
        trading_tabs.addTab(sell_widget, "提解賣出")
        trading_tabs.addTab(recording_widget, "錄音")
        trading_tabs.addTab(vip_list_widget, "VIP名單")
        trading_tabs.addTab(vip_quote_widget, "特殊報價")
        trading_tabs.addTab(bargain_widget, "議價交易")
        trading_tabs.addTab(cbas_to_cb_widget, "實物履約")
        trading_tabs.addTab(expired_widget, "合約到期")
        trading_tabs.addTab(renewal_widget, "選擇權續期")
        trading_tabs.addTab(customer_widget, "常用客戶維護")
        
        trading_layout.addWidget(trading_tabs)
        trading_widget.setLayout(trading_layout)
        
        # 建立後台工作分頁
        backend_widget = QWidget()
        backend_layout = QHBoxLayout()  # 改為水平布局
        
        # 左邊按鈕區域
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        
        # ========== 結算相關群組 ==========
        group_settlement = QGroupBox("結算相關")
        group_settlement_layout = QVBoxLayout()
        group_settlement_layout.setAlignment(Qt.AlignHCenter)  # 群組內容居中對齊
        
        self.btn_send_bargain_trade = QPushButton("寄信：議價交易")
        self.btn_send_bargain_trade.setFixedSize(200, 50)
        self.btn_send_bargain_trade.setStyleSheet("QPushButton { text-align: left; padding-left: 10px; }")
        self.btn_send_bargain_trade.clicked.connect(lambda: send_bargain_trade_email(self.output_text_edit, self))
        
        self.btn_send_today_trade = QPushButton("寄信：今日交易筆數")
        self.btn_send_today_trade.setFixedSize(200, 50)
        self.btn_send_today_trade.setStyleSheet("QPushButton { text-align: left; padding-left: 10px; }")
        self.btn_send_today_trade.clicked.connect(lambda: send_today_trade_email(self.output_text_edit, self))
        
        group_settlement_layout.addWidget(self.btn_send_bargain_trade, alignment=Qt.AlignHCenter)
        group_settlement_layout.addWidget(self.btn_send_today_trade, alignment=Qt.AlignHCenter)
        group_settlement.setLayout(group_settlement_layout)
        
        # ========== 客戶相關群組 ==========
        group_customer = QGroupBox("客戶相關")
        group_customer_layout = QVBoxLayout()
        group_customer_layout.setAlignment(Qt.AlignHCenter)  # 群組內容居中對齊
        
        self.btn_generate_today_detail = QPushButton("產檔：今日交易明細")
        self.btn_generate_today_detail.setFixedSize(200, 50)
        self.btn_generate_today_detail.setStyleSheet("QPushButton { text-align: left; padding-left: 10px; }")
        self.btn_generate_today_detail.clicked.connect(lambda: generate_today_detail(self.output_text_edit, self))
        
        self.btn_generate_trade_confirmation = QPushButton("產檔：交易確認書")
        self.btn_generate_trade_confirmation.setFixedSize(200, 50)
        self.btn_generate_trade_confirmation.setStyleSheet("QPushButton { text-align: left; padding-left: 10px; }")
        self.btn_generate_trade_confirmation.clicked.connect(lambda: generate_trade_confirmation(self.output_text_edit, self))
        
        # 分隔線
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.HLine)
        separator1.setFrameShadow(QFrame.Sunken)
        
        # 第二條分隔線
        separator2 = QFrame()
        separator2.setFrameShape(QFrame.HLine)
        separator2.setFrameShadow(QFrame.Sunken)
        
        self.btn_send_customer_detail = QPushButton("寄信：客戶當日成交明細")
        self.btn_send_customer_detail.setFixedSize(200, 50)
        self.btn_send_customer_detail.setStyleSheet("QPushButton { text-align: left; padding-left: 10px; }")
        self.btn_send_customer_detail.clicked.connect(lambda: send_customer_detail_email(self.output_text_edit, self))
        
        self.btn_send_customer_positions = QPushButton("寄信：客戶部位表")
        self.btn_send_customer_positions.setFixedSize(200, 50)
        self.btn_send_customer_positions.setStyleSheet("QPushButton { text-align: left; padding-left: 10px; }")
        self.btn_send_customer_positions.clicked.connect(lambda: send_customer_positions_email(self.output_text_edit, self))
        
        group_customer_layout.addWidget(self.btn_generate_today_detail, alignment=Qt.AlignHCenter)
        group_customer_layout.addWidget(self.btn_generate_trade_confirmation, alignment=Qt.AlignHCenter)
        group_customer_layout.addWidget(separator1)
        group_customer_layout.addWidget(separator2)
        group_customer_layout.addWidget(self.btn_send_customer_detail, alignment=Qt.AlignHCenter)
        group_customer_layout.addWidget(self.btn_send_customer_positions, alignment=Qt.AlignHCenter)
        group_customer.setLayout(group_customer_layout)
        
        # ========== CBAS相關群組 ==========
        group_cbas = QGroupBox("CBAS相關")
        group_cbas_layout = QVBoxLayout()
        group_cbas_layout.setAlignment(Qt.AlignHCenter)  # 群組內容居中對齊
        
        self.btn_send_control_table = QPushButton("產檔：控管表")
        self.btn_send_control_table.setFixedSize(200, 50)
        self.btn_send_control_table.setStyleSheet("QPushButton { text-align: left; padding-left: 10px; }")
        self.btn_send_control_table.clicked.connect(lambda: send_control_table_email(self.output_text_edit, self))
        
        group_cbas_layout.addWidget(self.btn_send_control_table, alignment=Qt.AlignHCenter)
        group_cbas.setLayout(group_cbas_layout)
        
        # 清除輸出按鈕（不屬於任何群組，放在最下面，一半大小）
        self.btn_clear_output = QPushButton("清除輸出")
        self.btn_clear_output.setFixedSize(100, 25)  # 一半大小：100x25
        self.btn_clear_output.setStyleSheet("QPushButton { text-align: left; padding-left: 10px; }")
        self.btn_clear_output.clicked.connect(lambda: clear_output_window(self.output_text_edit, self))
        
        # 添加到左邊布局，群組之間添加間距
        left_layout.addWidget(group_settlement)
        left_layout.addSpacing(15)  # 群組間距
        left_layout.addWidget(group_customer)
        left_layout.addSpacing(15)  # 群組間距
        left_layout.addWidget(group_cbas)
        left_layout.addStretch()  # 添加彈性空間
        left_layout.addWidget(self.btn_clear_output, alignment=Qt.AlignHCenter)  # 清除輸出按鈕居中對齊
        
        left_panel.setLayout(left_layout)
        left_panel.setFixedWidth(280)  # 固定左邊面板寬度
        
        # 右邊輸出視窗區域
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        
        # 輸出視窗標題
        output_label = QLabel("執行輸出:")
        output_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        right_layout.addWidget(output_label)
        
        # 建立輸出文字編輯器
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #f0f0f0;
                border: 1px solid #ccc;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
            }
        """)
        right_layout.addWidget(self.output_text_edit)
        
        right_panel.setLayout(right_layout)
        
        # 將左右面板添加到主布局
        backend_layout.addWidget(left_panel)
        backend_layout.addWidget(right_panel)
        backend_widget.setLayout(backend_layout)
        
        # 添加分頁到主tabs
        self.tabs.addTab(trading_widget, "交易處理")
        self.tabs.addTab(backend_widget, "後台工作")
        
        main_layout.addWidget(self.tabs)
        self.setLayout(main_layout)
        
        # 初始化報價表窗口為None
        self.quote_window = None
        
        self.loading_vips()
        # 讀取常用客戶資料
        self.load_customer_list()
        # 初始化報價表資料，在程式啟動時讀取一次
        self.df_quote, self.duplicate_cb, self.df_cbinfo = load_quote()
        self.examin_quote_duplicate(self.duplicate_cb)
        
        
        
        # 在UI完全創建後設置搜尋功能
        setup_exercise_input_search(self.input_cus_id, self.input_cb_code, self.df_quote, self.get_customer_list)
        self.setup_renewal_input_search()
        
        # 為所有表格安裝鍵盤刪除事件過濾器
        self.setup_table_keyboard_delete()

#====================AP功能區====================

    def loading_vips(self):
        """載入VIP資料並更新UI表格"""
        from file_reader import load_vip_data, get_daily_bond_rate
        
        # 使用 file_reader 模組讀取VIP資料
        df_vip_list, df_vip_quote = load_vip_data()
        
        # 更新VIP名單表格
        self.table_vip_list.setRowCount(len(df_vip_list))
        for i, row in df_vip_list.iterrows():
            for j, col in enumerate(df_vip_list.columns):
                item = QTableWidgetItem(str(row[col]))
                self.table_vip_list.setItem(i, j, item)

        # 更新特殊報價資料表格
        self.table_vip_quote.setRowCount(len(df_vip_quote))
        for i, row in df_vip_quote.iterrows():
            for j, col in enumerate(df_vip_quote.columns):
                item = QTableWidgetItem(str(row[col]))
                self.table_vip_quote.setItem(i, j, item)
    
    def get_vip_quote(self, vip_id):
        # 使用 get_table_data 方法讀取表格中的所有資料
        df_vip_quote = self.get_table_data(self.table_vip_quote)
        
        # 如果有指定 vip_id，則過濾特定客戶的報價
        if vip_id and not df_vip_quote.empty:
            df_filtered = df_vip_quote[df_vip_quote['客戶ID'] == vip_id]
            return df_filtered
        else:
            # 如果沒有指定 vip_id，返回所有特殊報價資料
            return df_vip_quote

    def add_row_specific(self, tab_name):
        """為指定分頁添加新列"""
        # 根據分頁名稱獲取對應的表格
        if tab_name == "新作買進":
            table = self.table_buy
        elif tab_name == "提解賣出":
            table = self.table_sell
        elif tab_name == "VIP名單":
            table = self.table_vip_list
        elif tab_name == "特殊報價":
            table = self.table_vip_quote
        elif tab_name == "議價交易":
            table = self.table_bargain
        elif tab_name == "實物履約":
            table = self.table_cbas_to_cb
        elif tab_name == "合約到期":
            table = self.table_expired
        elif tab_name == "選擇權續期":
            table = self.table_renewal_buy  # 默認新增到新作表格
        elif tab_name == "常用客戶維護":
            table = self.table_customer
        else:
            return
            
        row_count = table.rowCount()
        table.insertRow(row_count)
        
        # 為每個欄位設定預設值
        for col in range(table.columnCount()):
            item = QTableWidgetItem("")
            table.setItem(row_count, col, item)
        
        # 議價交易分頁的特殊處理
        if tab_name == "議價交易":
            # 自動產生單據編號（從10001開始）
            next_doc_number = self.get_next_bargain_doc_number()
            
            # 自動填入今天日期
            today = datetime.now().strftime('%Y%m%d')  # 格式: 20241225
            # 計算下一個工作日作為交割日期
            today_date = datetime.now()
            clearing_date_obj = next_business_day(today_date, 1)  # T+1 工作日
            clearing_date = clearing_date_obj.strftime('%Y%m%d')
            
            # 填入單據編號、今天日期、T+1、買/賣、交割日期
            table.setItem(row_count, 0, QTableWidgetItem(str(next_doc_number)))  # 單據編號
            table.setItem(row_count, 1, QTableWidgetItem(today))  # 成交日期
            table.setItem(row_count, 2, QTableWidgetItem("1"))  # T+?交割
            table.setItem(row_count, 3, QTableWidgetItem("買"))  # 買/賣（預設為買）
            table.setItem(row_count, 8, QTableWidgetItem("當日收盤價"))  # 參考價（預設為當日收盤價）
            
            # 將前9個欄位設定為淡黃色背景
            light_yellow = QColor(255, 255, 224)  # 淡黃色 RGB(255, 255, 224)
            for col in range(min(10, table.columnCount())):  # 前9個欄位或到欄位總數
                item = table.item(row_count, col)
                if item:
                    item.setBackground(light_yellow)

    def get_next_bargain_doc_number(self):
        """取得下一個議價交易單據編號"""
        table = self.table_bargain
        max_number = 10000  # 預設從10001開始
        
        # 檢查現有表格中的單據編號
        for row in range(table.rowCount()):
            item = table.item(row, 0)  # 第一欄是單據編號
            if item and item.text().strip():
                try:
                    number = int(item.text().strip())
                    if number > max_number:
                        max_number = number
                except ValueError:
                    continue
        
        return max_number + 1

    def get_customer_list(self):
        """獲取常用客戶列表"""
        customers = []
        # 檢查table_customer是否存在
        if not hasattr(self, 'table_customer'):
            return customers
        
        table = self.table_customer
        for row in range(table.rowCount()):
            customer_id_item = table.item(row, 0)
            customer_name_item = table.item(row, 1)
            if customer_id_item and customer_name_item:
                customer_id = customer_id_item.text().strip()
                customer_name = customer_name_item.text().strip()
                if customer_id and customer_name:
                    customers.append(f"{customer_id} - {customer_name}")
        return customers
    
    def setup_renewal_input_search(self):
        """設置選擇權續期輸入框的搜尋功能"""
        # 設置客戶ID搜尋功能
        customer_list = self.get_customer_list()
        for customer in customer_list:
            self.input_renewal_cus_id.addItem(customer)
        
        # 設置CB代號搜尋功能
        if hasattr(self, 'df_quote') and self.df_quote is not None:
            for _, row in self.df_quote.iterrows():
                cb_code = str(row.get('CB代號', '')).strip()
                cb_name = str(row.get('CB名稱', '')).strip()
                if cb_code and cb_name:
                    self.input_renewal_cb_code.addItem(f"{cb_code} - {cb_name}")
        
        # 暫時禁用搜尋功能以避免問題
        # self.input_renewal_cus_id.currentTextChanged.connect(self.filter_renewal_customer_items)
        # self.input_renewal_cb_code.currentTextChanged.connect(self.filter_renewal_cb_items)
    
    def setup_table_keyboard_delete(self):
        """為所有表格安裝鍵盤刪除事件過濾器（已改用TableWidgetWithDelete類，此方法保留以備用）"""
        pass
    
    def filter_renewal_customer_items(self, search_text):
        """過濾選擇權續期客戶ID項目"""
        # 如果正在更新中，跳過
        if hasattr(self, '_updating_renewal_customer') and self._updating_renewal_customer:
            return
        
        self._updating_renewal_customer = True
        
        try:
            # 保存當前選擇的項目
            current_text = self.input_renewal_cus_id.currentText()
            
            if not search_text:
                # 如果搜尋文字為空，顯示所有項目
                self.input_renewal_cus_id.clear()
                customer_list = self.get_customer_list()
                for customer in customer_list:
                    self.input_renewal_cus_id.addItem(customer)
            else:
                # 清空並重新載入符合條件的項目
                self.input_renewal_cus_id.clear()
                search_text = search_text.strip().upper()
                
                customer_list = self.get_customer_list()
                for customer in customer_list:
                    if search_text in customer.upper():
                        self.input_renewal_cus_id.addItem(customer)
            
            # 嘗試恢復之前的選擇
            if current_text and self.input_renewal_cus_id.findText(current_text) >= 0:
                self.input_renewal_cus_id.setCurrentText(current_text)
                
        except Exception as e:
            print(f"過濾選擇權續期客戶ID時發生錯誤：{e}")
        finally:
            self._updating_renewal_customer = False
    
    def filter_renewal_cb_items(self, search_text):
        """過濾選擇權續期CB代號項目"""
        # 如果正在更新中，跳過
        if hasattr(self, '_updating_renewal_cb') and self._updating_renewal_cb:
            return
        
        self._updating_renewal_cb = True
        
        try:
            # 保存當前選擇的項目
            current_text = self.input_renewal_cb_code.currentText()
            
            if not search_text:
                # 如果搜尋文字為空，顯示所有項目
                self.input_renewal_cb_code.clear()
                if hasattr(self, 'df_quote') and self.df_quote is not None:
                    for _, row in self.df_quote.iterrows():
                        cb_code = str(row.get('CB代號', '')).strip()
                        cb_name = str(row.get('CB名稱', '')).strip()
                        if cb_code and cb_name:
                            self.input_renewal_cb_code.addItem(f"{cb_code} - {cb_name}")
            else:
                # 清空並重新載入符合條件的項目
                self.input_renewal_cb_code.clear()
                search_text = search_text.strip().upper()
                
                if hasattr(self, 'df_quote') and self.df_quote is not None:
                    for _, row in self.df_quote.iterrows():
                        cb_code = str(row.get('CB代號', '')).strip()
                        cb_name = str(row.get('CB名稱', '')).strip()
                        
                        # 搜尋CB代號或CB名稱
                        if (search_text in cb_code.upper() or 
                            search_text in cb_name.upper()):
                            self.input_renewal_cb_code.addItem(f"{cb_code} - {cb_name}")
            
            # 嘗試恢復之前的選擇
            if current_text and self.input_renewal_cb_code.findText(current_text) >= 0:
                self.input_renewal_cb_code.setCurrentText(current_text)
                
        except Exception as e:
            print(f"過濾選擇權續期CB代號時發生錯誤：{e}")
        finally:
            self._updating_renewal_cb = False

    def delete_row_specific(self, tab_name):
        """刪除指定分頁的選定行"""
        # 根據分頁名稱獲取對應的表格
        if tab_name == "新作買進":
            table = self.table_buy
        elif tab_name == "提解賣出":
            table = self.table_sell
        elif tab_name == "VIP名單":
            table = self.table_vip_list
        elif tab_name == "特殊報價":
            table = self.table_vip_quote
        elif tab_name == "議價交易":
            table = self.table_bargain
        elif tab_name == "實物履約":
            # 檢查哪個表格有選中的行，優先刪除有選中行的表格
            selected_result = self.table_exercise_result.selectedItems()
            selected_cbas = self.table_cbas_to_cb.selectedItems()
            
            if selected_result:
                table = self.table_exercise_result
            elif selected_cbas:
                table = self.table_cbas_to_cb
            else:
                QMessageBox.warning(self, "警告", "請先選擇要刪除的行！（實物履約分頁）")
                return
        elif tab_name == "合約到期":
            table = self.table_expired
        elif tab_name == "選擇權續期":
            # 檢查哪個表格有選中的行，優先刪除有選中行的表格
            selected_buy = self.table_renewal_buy.selectedItems()
            selected_sell = self.table_renewal_sell.selectedItems()
            selected_query = self.table_renewal_query.selectedItems()
            
            if selected_sell:
                table = self.table_renewal_sell
            elif selected_buy:
                table = self.table_renewal_buy
            elif selected_query:
                table = self.table_renewal_query
            else:
                QMessageBox.warning(self, "警告", "請先選擇要刪除的行！（選擇權續期分頁）")
                return
        elif tab_name == "常用客戶維護":
            table = self.table_customer
        else:
            return
        
        # 獲取選定的行
        selected_rows = set()
        selected_items = table.selectedItems()
        
        if not selected_items:
            QMessageBox.warning(self, "警告", f"請先選擇要刪除的行！（{tab_name}分頁）")
            return
        
        # 收集所有選定項目的行號
        for item in selected_items:
            selected_rows.add(item.row())
        
        # 確認刪除
        if len(selected_rows) == 1:
            reply = QMessageBox.question(self, "確認刪除", f"確定要刪除{tab_name}分頁第 {min(selected_rows)+1} 行嗎？",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        else:
            reply = QMessageBox.question(self, "確認刪除", f"確定要刪除{tab_name}分頁選定的 {len(selected_rows)} 行嗎？",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # 按照從高到低的順序刪除行（避免索引變化問題）
            for row in sorted(selected_rows, reverse=True):
                table.removeRow(row)
            
            QMessageBox.information(self, "刪除成功", f"已刪除{tab_name}分頁 {len(selected_rows)} 行資料")

    def generate_buy_upload_file(self):
        """產生新作買進上傳檔"""
        try:
            today = datetime.now().strftime('%Y%m%d')
            today_dir = rf"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\{today}\上傳備份"
            
            # 檢查並創建今日資料夾
            if not os.path.exists(today_dir):
                os.makedirs(today_dir)
            
            # 讀取表格資料
            df = self.get_table_data(self.table_buy)
            
            if df.empty:
                QMessageBox.warning(self, "警告", "新作買進分頁沒有資料！")
                return
            
            # 定義上傳檔欄位順序
            upload_columns = ['上傳序號', '交易類型', '客戶ID', '交易日', '生效日', '交割日期', '選擇權到期日', 'CB代號', '提前履約界限日', '提前履約賠償金', '賣回日', '賣回價', '成交張數', '履約利率%',
                           '標的波動率', '無風險利率','資金成本', '轉債面額', '成交均價', '權利金百元價', '選擇權型態', '選擇權買賣別', '報價方式', '短契約', '手續費(業務單位)', '手續費(營業員)', '交易員', '營業員', '成交金額', '單位權利金', '權利金總額', '錄音日期', '錄音時間', '錄音人員', '子帳號', '固定端契約編號', '長約附加條款', '價格事件']
            
            # 確保欄位順序正確
            missing_columns = [col for col in upload_columns if col not in df.columns]
            for col in missing_columns:
                df[col] = ''
            
            # 重新排列欄位順序
            df_upload = df[upload_columns]
            df_upload['成交均價'] = df_upload['成交均價'].apply(lambda x: format_number_to_11(x, 11))
            df_upload['單位權利金'] = df_upload['單位權利金'].apply(lambda x: format_number_to_11(x, 11))
            
            market_qty = len(df[df['來自'] == '盤面交易'])
            bargain_qty = len(df[df['來自'] == '議價交易'])
            renewal_qty = len(df[df['來自'] == '續期'])

            len_upload = len(df_upload)
            reply = QMessageBox.question(self, "資料確認", f"盤面交易: {market_qty}筆資料\n議價交易: {bargain_qty}筆資料\n續期: {renewal_qty}筆資料\n總共產出: {len_upload}筆資料，請確認是否正確。\n\n確認報價方式是否有'2'", 
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
            
            # 儲存檔案（不含標題行）
            upload_file_path = os.path.join(today_dir, f"新作上傳檔_{today}.csv")
            df_upload.to_csv(upload_file_path, index=False, encoding='utf-8-sig', header=False)
            df.to_excel(os.path.join(today_dir, f"新作上傳檔_{today}.xlsx"), index=False)
            
            QMessageBox.information(self, "產生成功", f"新作買進上傳檔已產生：\n{upload_file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "產生失敗", f"發生錯誤：{e}")

    def generate_sell_upload_file(self):
        """產生提解賣出上傳檔"""
        try:
            today = datetime.now().strftime('%Y%m%d')
            today_dir = rf"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\{today}\上傳備份"
            
            # 檢查並創建今日資料夾
            if not os.path.exists(today_dir):
                os.makedirs(today_dir)
            
            # 讀取表格資料
            df = self.get_table_data(self.table_sell)
            
            if df.empty:
                QMessageBox.warning(self, "警告", "提解賣出分頁沒有資料！")
                return
            
            # 定義上傳檔欄位順序
            upload_columns = ['原單契約編號', '客戶ID', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
                           '履約張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日', '提前履約賠償金', '履約價', '選擇權交割單價', '交割總金額']
            
            # 確保欄位順序正確
            missing_columns = [col for col in upload_columns if col not in df.columns]
            for col in missing_columns:
                df[col] = ''
            
            # 重新排列欄位順序
            df_upload = df[upload_columns]

            df_upload['成交均價'] = df_upload['成交均價'].apply(lambda x: format_number_to_11(x, 11))
            df_upload['選擇權交割單價'] = df_upload['選擇權交割單價'].apply(lambda x: format_number_to_11(x, 11))


            market_qty = len(df[df['來自'] == '盤面交易'])
            bargain_qty = len(df[df['來自'] == '議價交易'])
            expired_qty = len(df[df['來自'] == '到期'])
            renewal_qty = len(df[df['來自'] == '續期'])
            cbas_to_cb_qty = len(df[df['來自'] == '實物履約'])


            len_upload = len(df_upload)
            reply = QMessageBox.question(self, "資料確認", f"盤面交易: {market_qty}筆資料\n議價交易: {bargain_qty}筆資料\n到期: {expired_qty}筆資料\n續期: {renewal_qty}筆資料\n實物履約: {cbas_to_cb_qty}筆資料\n總共產出: {len_upload}筆資料，請確認是否正確。\n\n確認報價方式是否有'2'", 
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
            
            # 上傳檔案路徑
            upload_file_path = os.path.join(today_dir, f"解約上傳檔_{today}.csv")
            
            # 儲存檔案（不含標題行）
            df_upload.to_csv(upload_file_path, index=False, encoding='utf-8-sig', header=False)
            df.to_excel(os.path.join(today_dir, f"解約上傳檔_{today}.xlsx"), index=False)
            
            QMessageBox.information(self, "產生成功", f"提解賣出上傳檔已產生：\n{upload_file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "產生失敗", f"發生錯誤：{e}")

    def open_upload_folder(self):
        """打開上傳檔資料夾"""
        try:
            folder_path = r'\\10.72.228.112\cbas業務公用區\CBAS上傳檔'
            if os.path.exists(folder_path):
                os.startfile(folder_path)
            else:
                QMessageBox.warning(self, "警告", f"資料夾不存在：\n{folder_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"打開資料夾時發生錯誤：{e}")

    def generate_recording_file(self):
        """產生錄音檔"""
        try:
            df_rec = self.get_table_data(self.table_recording)
            if df_rec.empty:
                QMessageBox.warning(self, "警告", "錄音表格沒有資料！")
                return
            
            # 產生錄音檔
            df_rec.insert(0, 'Date', datetime.now().strftime('%Y%m%d'))
            if os.path.exists(r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\錄音檔.xlsx"):
                df_rec_existing = pd.read_excel(r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\錄音檔.xlsx")
            else:
                df_rec_existing = pd.DataFrame(columns=df_rec.columns)
            
            df_rec_final = pd.concat([df_rec_existing, df_rec])
            df_rec_final.to_excel(r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\錄音檔.xlsx", index=False)
            
            # 將錄音時間填回買進和賣出表格
            self.fill_recording_time_back(df_rec)
            
            QMessageBox.information(self, "產生成功", "錄音檔已產：CBAS_Trading_Maker\錄音檔.xlsx")
            
        except Exception as e:
            QMessageBox.critical(self, "產生失敗", f"發生錯誤：{e}")
            import traceback
            traceback.print_exc()
    
    def fill_recording_time_back(self, df_recording):
        """將錄音時間填回買進和賣出表格
        使用字符串拼接的 match 欄位進行批量匹配，效能更佳
        """
        try:
            # 找到錄音時間欄位的索引
            buy_columns = [self.table_buy.horizontalHeaderItem(i).text() for i in range(self.table_buy.columnCount())]
            sell_columns = [self.table_sell.horizontalHeaderItem(i).text() for i in range(self.table_sell.columnCount())]
            
            buy_recording_time_col_idx = buy_columns.index('錄音時間') if '錄音時間' in buy_columns else -1
            sell_recording_time_col_idx = sell_columns.index('錄音時間') if '錄音時間' in sell_columns else -1
            
            if buy_recording_time_col_idx == -1 and sell_recording_time_col_idx == -1:
                return  # 如果兩個表格都沒有錄音時間欄位，直接返回
            
            # 處理買進表格：使用 match 欄位進行批量匹配
            if buy_recording_time_col_idx != -1:
                # 讀取買進表格資料
                df_buy = self.get_table_data(self.table_buy)
                
                if not df_buy.empty:
                    # 篩選錄音表格中買進張數不為空的記錄
                    df_rec_buy = df_recording[df_recording['買進張數'].astype(str).str.strip() != ''].copy()
                    
                    if not df_rec_buy.empty:
                        # 標準化數據類型
                        df_buy['客戶ID'] = df_buy['客戶ID'].astype(str).str.strip()
                        df_buy['CB代號'] = df_buy['CB代號'].astype(str).str.strip()
                        df_buy['成交張數'] = df_buy['成交張數'].astype(str).str.strip()
                        df_buy['成交均價'] = pd.to_numeric(df_buy['成交均價'], errors='coerce')
                        
                        df_rec_buy['客戶ID'] = df_rec_buy['客戶ID'].astype(str).str.strip()
                        df_rec_buy['CB代號'] = df_rec_buy['CB代號'].astype(str).str.strip()
                        df_rec_buy['買進張數'] = df_rec_buy['買進張數'].astype(str).str.strip()
                        df_rec_buy['成交均價'] = pd.to_numeric(df_rec_buy['成交均價'], errors='coerce')
                        
                        # 創建 match 欄位：客戶ID+CB代號+成交張數+成交均價
                        df_buy['match'] = (df_buy['客戶ID'] + '|' + df_buy['CB代號'] + '|' + 
                                          df_buy['成交張數'] + '|' + df_buy['成交均價'].astype(str))
                        df_rec_buy['match'] = (df_rec_buy['客戶ID'] + '|' + df_rec_buy['CB代號'] + '|' + 
                                               df_rec_buy['買進張數'] + '|' + df_rec_buy['成交均價'].astype(str))
                        
                        # 建立 match 到錄音時間的映射字典
                        match_to_time = dict(zip(df_rec_buy['match'], df_rec_buy['錄音時間'].astype(str).str.strip()))
                        
                        # 批量匹配並更新
                        df_buy['錄音時間_matched'] = df_buy['match'].map(match_to_time).fillna('')
                        
                        # 更新表格顯示
                        for i in range(min(len(df_buy), self.table_buy.rowCount())):
                            recording_time = str(df_buy.iloc[i]['錄音時間_matched']).strip()
                            if recording_time:
                                item = self.table_buy.item(i, buy_recording_time_col_idx)
                                if item:
                                    item.setText(recording_time)
                                else:
                                    item = QTableWidgetItem(recording_time)
                                    self.table_buy.setItem(i, buy_recording_time_col_idx, item)
            
            # 處理賣出表格：使用 match 欄位進行批量匹配
            if sell_recording_time_col_idx != -1:
                # 讀取賣出表格資料
                df_sell = self.get_table_data(self.table_sell)
                
                if not df_sell.empty:
                    # 篩選錄音表格中賣出張數不為空的記錄
                    df_rec_sell = df_recording[df_recording['賣出張數'].astype(str).str.strip() != ''].copy()
                    
                    if not df_rec_sell.empty:
                        # 標準化數據類型
                        df_sell['客戶ID'] = df_sell['客戶ID'].astype(str).str.strip()
                        df_sell['CB代號'] = df_sell['CB代號'].astype(str).str.strip()
                        df_sell['履約張數'] = df_sell['履約張數'].astype(str).str.strip()
                        df_sell['成交均價'] = pd.to_numeric(df_sell['成交均價'], errors='coerce')
                        
                        df_rec_sell['客戶ID'] = df_rec_sell['客戶ID'].astype(str).str.strip()
                        df_rec_sell['CB代號'] = df_rec_sell['CB代號'].astype(str).str.strip()
                        df_rec_sell['賣出張數'] = df_rec_sell['賣出張數'].astype(str).str.strip()
                        df_rec_sell['成交均價'] = pd.to_numeric(df_rec_sell['成交均價'], errors='coerce')
                        
                        # 創建 match 欄位：客戶ID+CB代號+履約張數+成交均價
                        df_sell['match'] = (df_sell['客戶ID'] + '|' + df_sell['CB代號'] + '|' + 
                                           df_sell['履約張數'] + '|' + df_sell['成交均價'].astype(str))
                        df_rec_sell['match'] = (df_rec_sell['客戶ID'] + '|' + df_rec_sell['CB代號'] + '|' + 
                                               df_rec_sell['賣出張數'] + '|' + df_rec_sell['成交均價'].astype(str))
                        
                        # 建立 match 到錄音時間的映射字典
                        match_to_time = dict(zip(df_rec_sell['match'], df_rec_sell['錄音時間'].astype(str).str.strip()))
                        
                        # 批量匹配並更新
                        df_sell['錄音時間_matched'] = df_sell['match'].map(match_to_time).fillna('')
                        
                        # 更新表格顯示
                        for i in range(min(len(df_sell), self.table_sell.rowCount())):
                            recording_time = str(df_sell.iloc[i]['錄音時間_matched']).strip()
                            if recording_time:
                                item = self.table_sell.item(i, sell_recording_time_col_idx)
                                if item:
                                    item.setText(recording_time)
                                else:
                                    item = QTableWidgetItem(recording_time)
                                    self.table_sell.setItem(i, sell_recording_time_col_idx, item)
                                    
        except Exception as e:
            print(f"填回錄音時間時發生錯誤：{e}")
            import traceback
            traceback.print_exc()
    
    def refresh_recording_table(self):
        """刷新錄音表格資料"""
        try:
            # 讀取新作買進表格資料
            df_buy = self.get_table_data(self.table_buy)
            # 讀取提解賣出表格資料
            df_sell = self.get_table_data(self.table_sell)
            
            # 篩選錄音時間不為'E'的資料
            if not df_buy.empty:
                df_buy_filtered = df_buy[df_buy['錄音時間'].astype(str).str.strip() != 'E'].copy()
            else:
                df_buy_filtered = pd.DataFrame()
            
            if not df_sell.empty:
                df_sell_filtered = df_sell[df_sell['錄音時間'].astype(str).str.strip() != 'E'].copy()
            else:
                df_sell_filtered = pd.DataFrame()
            
            # 如果兩個表格都沒有資料
            if df_buy_filtered.empty and df_sell_filtered.empty:
                self.table_recording.setRowCount(0)
                QMessageBox.information(self, "提示", "沒有符合條件的資料（錄音時間不為E）！")
                return
            
            # 準備新作買進的資料
            buy_records = []
            if not df_buy_filtered.empty:
                for _, row in df_buy_filtered.iterrows():
                    buy_records.append({
                        '客戶ID': str(row.get('客戶ID', '')).strip(),
                        '客戶名稱': str(row.get('客戶名稱', '')).strip(),
                        'CB名稱': str(row.get('CB名稱', '')).strip(),
                        'CB代號': str(row.get('CB代號', '')).strip(),
                        '買進張數': str(row.get('成交張數', '')).strip(),
                        '賣出張數': '',  # 買進沒有賣出張數
                        '成交均價': str(row.get('成交均價', '')).strip(),
                    })
            
            # 準備提解賣出的資料
            sell_records = []
            if not df_sell_filtered.empty:
                for _, row in df_sell_filtered.iterrows():
                    sell_records.append({
                        '客戶ID': str(row.get('客戶ID', '')).strip(),
                        '客戶名稱': str(row.get('客戶名稱', '')).strip(),
                        'CB名稱': str(row.get('CB名稱', '')).strip(),
                        'CB代號': str(row.get('CB代號', '')).strip(),
                        '買進張數': '',  # 賣出沒有買進張數
                        '賣出張數': str(row.get('履約張數', '')).strip(),
                        '成交均價': str(row.get('成交均價', '')).strip(),
                    })
            
            # 合併資料
            all_records = buy_records + sell_records
            
            if not all_records:
                self.table_recording.setRowCount(0)
                QMessageBox.information(self, "提示", "沒有符合條件的資料！")
                return
            
            # 轉換為DataFrame
            df_recording = pd.DataFrame(all_records)
            
            # 按客戶ID排序
            df_recording = df_recording.sort_values('客戶ID').reset_index(drop=True)
            
            # 獲取所有客戶ID並取得CELLPHONE
            cusid_list = df_recording['客戶ID'].unique().tolist()
            if cusid_list:
                cus_info = get_customer_bank_and_email(cusid_list)
                trust_info = get_trust_info(cusid_list)
                # 合併CELLPHONE欄位
                df_recording = df_recording.merge(
                    cus_info[['CUSID', 'CELLPHONE']], 
                    left_on='客戶ID', 
                    right_on='CUSID', 
                    how='left'
                )
                df_recording = df_recording.merge(
                    trust_info[['CUSID', 'TRUSTNM', 'TRUSTTEL']], 
                    left_on='客戶ID', 
                    right_on='CUSID', 
                    how='left'
                )
                # 移除臨時的CUSID欄位
                if 'CUSID' in df_recording.columns:
                    df_recording = df_recording.drop(columns=['CUSID'])
            else:
                # 如果沒有客戶ID，添加空的CELLPHONE欄位
                df_recording['CELLPHONE'] = ''
                df_recording['TRUSTNM'] = ''
                df_recording['TRUSTTEL'] = ''

            df_recording['授權人'] = df_recording['TRUSTNM']
            df_recording['授權人電話'] = df_recording['TRUSTTEL']
            df_recording = df_recording.drop(columns=['TRUSTNM', 'TRUSTTEL'])
            # 添加錄音時間欄位
            df_recording['錄音時間'] = ''
            # 添加錄音人員欄位，默認值為"蔡睿"
            df_recording['錄音人員'] = '蔡睿'
            
            # 重新排列欄位順序
            df_recording = df_recording[['客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '買進張數', '賣出張數', '成交均價', 'CELLPHONE', '授權人', '授權人電話', '錄音時間', '錄音人員']]
            
            # 更新表格顯示
            self.update_recording_table(df_recording)
            
        except Exception as e:
            QMessageBox.critical(self, "刷新失敗", f"發生錯誤：{e}")
            import traceback
            traceback.print_exc()
    
    def update_recording_table(self, df_recording):
        """更新錄音表格顯示"""
        try:
            self.table_recording.setRowCount(len(df_recording))
            
            # 獲取表格的列名
            recording_columns = [self.table_recording.horizontalHeaderItem(i).text() for i in range(self.table_recording.columnCount())]
            
            for i, row in df_recording.iterrows():
                for j, col_name in enumerate(recording_columns):
                    # 如果DataFrame中有這個列，使用DataFrame的值；否則使用空字串
                    if col_name in df_recording.columns:
                        value = row[col_name] if pd.notna(row[col_name]) else ''
                    else:
                        value = ''
                    
                    # 如果是錄音人員列且值為空，設置默認值為"蔡睿"
                    if col_name == '錄音人員' and (not value or value == ''):
                        value = '蔡睿'
                    
                    item = QTableWidgetItem(str(value))
                    self.table_recording.setItem(i, j, item)
        except Exception as e:
            print(f"更新錄音表格時發生錯誤：{e}")
            import traceback
            traceback.print_exc()

    def save_vip_list(self):
        """儲存VIP名單"""
        try:
            # 讀取表格資料
            df = self.get_table_data(self.table_vip_list)
            
            if df.empty:
                QMessageBox.warning(self, "警告", "VIP名單分頁沒有資料！")
                return
            
            # VIP名單檔案路徑
            vip_list_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\VIP_List.csv"
            
            # 儲存檔案
            df.to_csv(vip_list_path, index=False, encoding='utf-8-sig', header=True)
            
            QMessageBox.information(self, "儲存成功", f"VIP名單已儲存至：\n{vip_list_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "儲存失敗", f"發生錯誤：{e}")

    def save_vip_quote(self):
        """儲存特殊報價"""
        try:
            # 讀取表格資料
            df = self.get_table_data(self.table_vip_quote)
            
            if df.empty:
                QMessageBox.warning(self, "警告", "特殊報價分頁沒有資料！")
                return
            
            # 特殊報價檔案路徑
            vip_quote_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\VIP_Quote.csv"
            
            # 儲存檔案
            df.to_csv(vip_quote_path, index=False, encoding='utf-8-sig', header=True)
            
            QMessageBox.information(self, "儲存成功", f"特殊報價已儲存至：\n{vip_quote_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "儲存失敗", f"發生錯誤：{e}")
    
    def save_customer_list(self):
        """儲存常用客戶資料"""
        try:
            # 讀取表格資料
            df = self.get_table_data(self.table_customer)
            
            if df.empty:
                QMessageBox.warning(self, "警告", "常用客戶維護分頁沒有資料！")
                return
            
            # 常用客戶檔案路徑
            customer_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\Customer_List.csv"
            
            # 儲存檔案
            df.to_csv(customer_path, index=False, encoding='utf-8-sig', header=True)
            
            QMessageBox.information(self, "儲存成功", f"常用客戶資料已儲存至：\n{customer_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "儲存失敗", f"發生錯誤：{e}")
    
    def load_customer_list(self):
        """讀取常用客戶資料"""
        try:
            customer_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\Customer_List.csv"
            
            # 嘗試多種編碼讀取
            df_customer = None
            for encoding in ['utf-8-sig', 'utf-8', 'big5', 'gbk', 'cp950']:
                try:
                    df_customer = pd.read_csv(customer_path, encoding=encoding)
                    break
                except (UnicodeDecodeError, FileNotFoundError):
                    continue
            
            # 如果檔案不存在，創建空的 DataFrame
            if df_customer is None:
                print(f"無法讀取 {customer_path}，將創建空的常用客戶名單")
                df_customer = pd.DataFrame(columns=['客戶ID', '客戶名稱'])
            
            # 更新常用客戶表格
            self.table_customer.setRowCount(len(df_customer))
            for i, row in df_customer.iterrows():
                for j, col in enumerate(df_customer.columns):
                    item = QTableWidgetItem(str(row[col]))
                    self.table_customer.setItem(i, j, item)
            
            print("常用客戶資料已載入完成")
            
        except Exception as e:
            print(f"讀取常用客戶資料錯誤：{e}")
            # 如果讀取失敗，創建空的表格
            self.table_customer.setRowCount(0)

    def delete_row(self):
        """刪除選定的行（保留原有方法以相容性）"""
        current_tab = self.tabs.currentIndex()
        tab_name = self.tab_names[current_tab]
        self.delete_row_specific(tab_name)

    def temp_save_all(self):
        """暫存所有交易處理分頁的資料"""
        try:
            # 確保temp資料夾存在
            temp_dir = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\temp"
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)
            
            # 定義所有表格、所屬分頁(tab_name)與表格名稱(table_name)
            # 依需求：暫存時不包含 VIP名單 與 特殊報價
            tables_data = [
                (self.table_buy, "新作買進", "table_buy"),
                (self.table_sell, "提解賣出", "table_sell"),
                (self.table_bargain, "議價交易", "table_bargain"),
                # 實物履約包含兩個表
                (self.table_exercise_result, "實物履約", "table_exercise_result"),
                (self.table_cbas_to_cb, "實物履約", "table_cbas_to_cb"),
                (self.table_expired, "合約到期", "table_expired"),
                # 選擇權續期包含三個表
                (self.table_renewal_query, "選擇權續期", "table_renewal_query"),
                (self.table_renewal_buy, "選擇權續期", "table_renewal_buy"),
                (self.table_renewal_sell, "選擇權續期", "table_renewal_sell"),
            ]
            
            all_data = []
            
            # 收集所有表格的資料
            for table, tab_name, table_name in tables_data:
                df = self.get_table_data(table)
                if not df.empty:
                    # 為每行資料添加識別欄位（tab_name, table_name, index_row 放在最前面）
                    for idx, row in df.iterrows():
                        row_data = row.to_dict()
                        row_data = {**{"tab_name": tab_name, "table_name": table_name, "index_row": idx}, **row_data}
                        all_data.append(row_data)
            
            if not all_data:
                QMessageBox.warning(self, "警告", "沒有資料可以暫存！")
                return
            
            # 轉換為DataFrame並確保欄位順序：tab_name, table_name, index_row 置前
            df_all = pd.DataFrame(all_data)
            lead_cols = [c for c in ["tab_name", "table_name", "index_row"] if c in df_all.columns]
            other_cols = [c for c in df_all.columns if c not in lead_cols]
            df_all = df_all[lead_cols + other_cols]
            
            temp_file_path = os.path.join(temp_dir, "temp.csv")
            df_all.to_csv(temp_file_path, index=False, encoding='utf-8-sig')
            
            QMessageBox.information(self, "暫存成功", f"所有資料已暫存至：\n{temp_file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "暫存失敗", f"發生錯誤：{e}")
    
    def temp_load_all(self):
        """讀取暫存的所有資料"""
        try:
            temp_file_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\temp\temp.csv"
            
            if not os.path.exists(temp_file_path):
                QMessageBox.warning(self, "警告", "暫存檔案不存在！")
                return
            
            # 讀取暫存檔案
            # 將 'nan' 字串與空白視為缺值，後續顯示為空
            df_all = pd.read_csv(temp_file_path, encoding='utf-8-sig', keep_default_na=True, na_values=['nan', 'NaN', 'None', ''])
            # 取得議價交易表格的欄位名稱
            tables = {
                "table_buy": self.table_buy,
                "table_sell": self.table_sell,
                "table_bargain": self.table_bargain,
                "table_exercise_result": self.table_exercise_result,
                "table_cbas_to_cb": self.table_cbas_to_cb,
                "table_expired": self.table_expired,
                "table_renewal_query": self.table_renewal_query,
                "table_renewal_buy": self.table_renewal_buy,
                "table_renewal_sell": self.table_renewal_sell,
            }

            # 遍歷所有表格
            for table_name, table in tables.items():
                # 嘗試從 df_all 中過濾對應資料
                df = df_all[df_all['table_name'] == table_name]
                print(f"處理表格: {table_name}")
                if not df.empty:
                    # 抓該表格欄位名稱（來自 header）
                    table_columns = [
                        table.horizontalHeaderItem(i).text()
                        if table.horizontalHeaderItem(i) else f"Column{i}"
                        for i in range(table.columnCount())
                    ]

                    table.setRowCount(len(df))
                    df = df[table_columns].reset_index(drop=True)
                    print(df)
                    # 將資料填入 QTableWidget 中
                    for i, row in df.iterrows():
                        for j, col in enumerate(df.columns):
                            val = row[col]
                            # 格式化數字：如果是整數就去掉 .0，如果是小數保留
                            if pd.notna(val) and isinstance(val, (int, float)):
                                if val == int(val):  # 如果是整數
                                    text = str(int(val))
                                else:  # 如果是小數
                                    text = str(val)
                            else:
                                text = str(val) if pd.notna(val) else ""
                            
                            item = QTableWidgetItem(text)
                            # 可加背景色：item.setBackground(QColor(230, 255, 230))
                            table.setItem(i, j, item)


        except Exception as e:
            QMessageBox.critical(self, "讀取失敗", f"發生錯誤：{e}")
 
    def save_one(self):
        """儲存當前分頁資料（保留原有方法以相容性）"""
        current_tab = self.tabs.currentIndex()
        tab_name = self.tab_names[current_tab]
        
        if tab_name == "新作買進":
            self.generate_buy_upload_file()
        elif tab_name == "提解賣出":
            self.generate_sell_upload_file()
        elif tab_name == "VIP名單":
            self.save_vip_list()
        elif tab_name == "特殊報價":
            self.save_vip_quote()
        elif tab_name == "議價交易":
            # 議價交易保存為 CSV
            today = datetime.now().strftime('%Y%m%d')
            today_dir = rf"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\{today}"
            if not os.path.exists(today_dir):
                os.makedirs(today_dir)
            
            df = self.get_table_data(self.table_bargain)
            if not df.empty:
                file_path = os.path.join(today_dir, "議價交易.csv")
                df.to_csv(file_path, index=False, encoding='utf-8-sig', header=True)
                QMessageBox.information(self, "儲存成功", f"議價交易已儲存至：\n{file_path}")
            else:
                QMessageBox.warning(self, "警告", "議價交易分頁沒有資料！")
        elif tab_name == "實物履約":
            # 實物履約保存為 CSV
            today = datetime.now().strftime('%Y%m%d')
            today_dir = rf"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\{today}"
            if not os.path.exists(today_dir):
                os.makedirs(today_dir)
            
            df = self.get_table_data(self.table_cbas_to_cb)
            if not df.empty:
                file_path = os.path.join(today_dir, "實物履約.csv")
                df.to_csv(file_path, index=False, encoding='utf-8-sig', header=True)
                QMessageBox.information(self, "儲存成功", f"實物履約已儲存至：\n{file_path}")
            else:
                QMessageBox.warning(self, "警告", "實物履約分頁沒有資料！")
        elif tab_name == "合約到期":
            # 合約到期保存為 CSV
            today = datetime.now().strftime('%Y%m%d')
            today_dir = rf"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\{today}"
            if not os.path.exists(today_dir):
                os.makedirs(today_dir)
            
            df = self.get_table_data(self.table_expired)
            if not df.empty:
                file_path = os.path.join(today_dir, "合約到期.csv")
                df.to_csv(file_path, index=False, encoding='utf-8-sig', header=True)
                QMessageBox.information(self, "儲存成功", f"合約到期已儲存至：\n{file_path}")
            else:
                QMessageBox.warning(self, "警告", "合約到期分頁沒有資料！")
        elif tab_name == "選擇權續期":
            # 選擇權續期保存為 CSV
            today = datetime.now().strftime('%Y%m%d')
            today_dir = rf"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\{today}"
            if not os.path.exists(today_dir):
                os.makedirs(today_dir)
            
            df = self.get_table_data(self.table_renewal_buy)
            if not df.empty:
                file_path = os.path.join(today_dir, "選擇權續期.csv")
                df.to_csv(file_path, index=False, encoding='utf-8-sig', header=True)
                QMessageBox.information(self, "儲存成功", f"選擇權續期已儲存至：\n{file_path}")
            else:
                QMessageBox.warning(self, "警告", "選擇權續期分頁沒有資料！")

    def add_row(self):
        """新增列（保留原有方法以相容性）"""
        current_tab = self.tabs.currentIndex()
        tab_name = self.tab_names[current_tab]
        self.add_row_specific(tab_name)

    def generate_i_realized_file(self):
        """製作已實現損益"""
        df_sell = self.get_table_data(self.table_sell)
        if df_sell.empty:
            QMessageBox.warning(self, "警告", "賣出表格沒有資料！")
            return
        df_contracts = get_contracts_from_sell_table(df_sell)
        df_sell = df_sell.merge(df_contracts, on='原單契約編號', how='left')
        df_sell['已實現損益'] = round((df_sell['選擇權交割單價'].astype(float) - df_sell['原單位權利金'].astype(float)) * 1000 * df_sell['履約張數'].astype(int), 0)
        df_sell['選擇權買賣別'] = 'S'
        df_reorg = df_sell[['原單契約編號', '客戶ID', 'CB代號', '交易日期','交割日期', '解約類別', '履約方式', '履約張數', '成交均價', '履約利率',
                            '賣回日', '賣回價', '選擇權到期日', '提前履約賠償金', '履約價', '選擇權交割單價', '交割總金額', '報價方式', '選擇權買賣別',
                            '原單位權利金', '已實現損益']]

        tday = datetime.now().strftime('%Y%m%d')
        df_reorg.to_excel(rf'\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\{tday}\i已實_CBAS契約檔_.xlsx', index=False, header=True)
        QMessageBox.information(self, "i已實儲存成功", "i已實儲存成功")

    def get_monitor_fill_sum_amt(self):
        df_monitor_fill = get_631_Monitor_Fill()
        sum_buy_amt = int(df_monitor_fill[~df_monitor_fill['客戶代碼'].isin(['23218183', ''])]['買進金額'].sum())
        sum_buy_qty = int(df_monitor_fill[~df_monitor_fill['客戶代碼'].isin(['23218183', ''])]['買進股數'].sum() / 1000)
        sum_sell_amt = int(df_monitor_fill[~df_monitor_fill['客戶代碼'].isin(['23218183', ''])]['賣出金額'].sum())
        sum_sell_qty = int(df_monitor_fill[~df_monitor_fill['客戶代碼'].isin(['23218183', ''])]['賣出股數'].sum() / 1000)
        print(f"CB自營系統買進: {sum_buy_amt:,} 元, {sum_buy_qty} 張")
        print(f"CB自營系統賣出: {sum_sell_amt:,} 元, {sum_sell_qty} 張")

        return sum_buy_amt, sum_buy_qty, sum_sell_amt, sum_sell_qty

    def check_buy_table_with_quote(self):
        
        df_buy = self.get_table_data(self.table_buy)
        if df_buy.empty:
            QMessageBox.warning(self, "警告", "新作買進分頁沒有資料！")
            return
        
        df_buy['權利金百元價'] = pd.to_numeric(df_buy['權利金百元價'], errors='coerce')
        df_buy['手續費(業務單位)'] = pd.to_numeric(df_buy['手續費(業務單位)'], errors='coerce')
        
        # 併上報價表（只取用必要欄位）
        quote_cols = ['CB代號', '百元報價', '低百元報價']
        df_quote_use = self.df_quote[quote_cols].rename(columns={'百元報價': '百元報價_quote', '低百元報價': '低百元報價_quote'})
        df_chk = df_buy.merge(df_quote_use, on='CB代號', how='left')
        
        # 手續費對應 offset 規則
        offset_map = {110: 0.04, 150: 0.00, 100: 0.05, 60: 0.09}
        df_chk['offset'] = df_chk['手續費(業務單位)'].map(offset_map)
        
        # 計算預期百元報價
        df_chk['預期價1'] = df_chk['百元報價_quote'] - df_chk['offset']
        df_chk['預期價2'] = df_chk['低百元報價_quote'] - df_chk['offset']
        
        mask = df_chk['手續費(業務單位)'] == 150
        # 預期價1 = 百元報價（不折）
        df_chk.loc[mask, '預期價1'] = df_chk.loc[mask, '百元報價_quote'].values

        # 預期價2 = 低百元報價（不折）
        df_chk.loc[mask, '預期價2'] = df_chk.loc[mask, '低百元報價_quote'].values
        # 允許小數誤差（四捨五入到 2 位小數再比對）
        def eq2(a, b):
            return (pd.notna(a) & pd.notna(b) & (np.round(a, 2) == np.round(b, 2)))
        
        match1 = eq2(df_chk['權利金百元價'], df_chk['預期價1'])
        match2 = eq2(df_chk['權利金百元價'], df_chk['預期價2'])
        df_chk['是否符合'] = match1 | match2
        
        # 依結果在表格上著色：符合=綠底，不符合=淺紅底
        col_idx_price = None
        for j in range(self.table_buy.columnCount()):
            header_item = self.table_buy.horizontalHeaderItem(j)
            if header_item and header_item.text() == '權利金百元價':
                col_idx_price = j
                break
        if col_idx_price is not None:
            for i in range(min(self.table_buy.rowCount(), len(df_chk))):
                item = self.table_buy.item(i, col_idx_price)
                if item is None:
                    # 若該格沒有項目，補上一個（避免 setBackground 失效）
                    val = df_chk.iloc[i]['權利金百元價']
                    item = QTableWidgetItem('' if pd.isna(val) else str(val))
                    self.table_buy.setItem(i, col_idx_price, item)
                ok = bool(df_chk.iloc[i]['是否符合']) if pd.notna(df_chk.iloc[i]['是否符合']) else False
                item.setBackground(QColor(230, 255, 230) if ok else QColor(255, 230, 230))
        
        # 整理不符合清單
        df_bad = df_chk[~df_chk['是否符合'].fillna(False)].copy()
        df_bad['說明'] = (
            '手續費=' + df_bad['手續費(業務單位)'].astype('Int64').astype(str) +
            ', 權利金百元價=' + df_bad['權利金百元價'].astype(str) +
            ', 預期=[' + df_bad['預期價1'].round(2).astype(str) + '/' + df_bad['預期價2'].round(2).astype(str) + ']'
        )
        
        total = len(df_chk)
        ok = int(df_chk['是否符合'].fillna(False).sum())
        ng = total - ok
        
        if ng == 0:
            QMessageBox.information(self, "檢查結果", f"全部符合（{ok}/{total}）")
        else:
            # 只展示前 20 筆不符合做為提示
            preview = '\n'.join(df_bad[['CB代號','說明']].head(20).apply(lambda r: f"{r['CB代號']}: {r['說明']}", axis=1))
            QMessageBox.warning(self, "檢查結果", f"不符合 {ng}/{total} 筆：\n{preview}")
        
        return

    def check_sell_table_with_qty(self):
        df_sell = self.get_table_data(self.table_sell)
        if df_sell.empty:
            QMessageBox.warning(self, "警告", "提解賣出分頁沒有資料！")
            return
         
        # 數值轉換
        df_sell['履約張數'] = pd.to_numeric(df_sell['履約張數'], errors='coerce').fillna(0)
        df_sell['成交均價'] = pd.to_numeric(df_sell['成交均價'], errors='coerce').fillna(0)
        df_sell['成交金額'] = round(df_sell['履約張數'] * df_sell['成交均價'] * 1000, 0)
        
        # 依來源彙總
        df_group = df_sell.groupby('來自', dropna=False).agg({'履約張數': 'sum', '成交金額': 'sum'}).reset_index()
        df_group['履約張數'] = df_group['履約張數'].astype(int)
        df_group['成交金額'] = df_group['成交金額'].astype(int)
        
        # 組訊息
        lines = []
        for _, r in df_group.iterrows():
            source = str(r['來自']) if pd.notna(r['來自']) else '未分類'
            qty = int(r['履約張數'])
            amt = int(r['成交金額'])
            lines.append(f"{source}: {qty} 張, {amt:,} 元")
        total_qty = int(df_group['履約張數'].sum())
        total_amt = int(df_group['成交金額'].sum())
        lines.append(f"總計: {total_qty} 張, {total_amt:,} 元")
        
        QMessageBox.information(self, "張數/總額確認", "\n".join(lines))

        sum_buy_amt, sum_buy_qty, sum_sell_amt, sum_sell_qty = self.get_monitor_fill_sum_amt()
        sellmatch_amt = int(df_group[df_group['來自'] == '盤面交易']['成交金額'].sum())
        sellmatch_qty = int(df_group[df_group['來自'] == '盤面交易']['履約張數'].sum())

        if sellmatch_amt != sum_sell_amt or sellmatch_qty != sum_sell_qty:
            QMessageBox.warning(self, "警告", f"盤面交易金額或張數與CB自營系統不符\n盤面交易: {sellmatch_amt:,} 元, {sellmatch_qty} 張\nCB自營系統: {sum_sell_amt:,} 元, {sum_sell_qty} 張")
        elif sellmatch_amt == sum_sell_amt and sellmatch_qty == sum_sell_qty:
            QMessageBox.information(self, "確認", "盤面交易金額或張數與CB自營系統相符")

    def check_buy_table_with_qty(self):
        df_buy = self.get_table_data(self.table_buy)
        if df_buy.empty:
            QMessageBox.warning(self, "警告", "新作買進分頁沒有資料！")
            return
        
        # 數值轉換
        df_buy['成交張數'] = pd.to_numeric(df_buy['成交張數'], errors='coerce').fillna(0)
        df_buy['成交均價'] = pd.to_numeric(df_buy['成交均價'], errors='coerce').fillna(0)
        df_buy['成交金額'] = round(df_buy['成交張數'] * df_buy['成交均價'] * 1000, 0)
        
        # 依來源彙總
        df_group = df_buy.groupby('來自', dropna=False).agg({'成交張數': 'sum', '成交金額': 'sum'}).reset_index()
        df_group['成交張數'] = df_group['成交張數'].astype(int)
        df_group['成交金額'] = df_group['成交金額'].astype(int)
        
        # 組訊息
        lines = []
        for _, r in df_group.iterrows():
            source = str(r['來自']) if pd.notna(r['來自']) else '未分類'
            qty = int(r['成交張數'])
            amt = int(r['成交金額'])
            lines.append(f"{source}: {qty} 張, {amt:,} 元")
        total_qty = int(df_group['成交張數'].sum())
        total_amt = int(df_group['成交金額'].sum())
        lines.append(f"總計: {total_qty} 張, {total_amt:,} 元")

        QMessageBox.information(self, "張數/總額確認", "\n".join(lines))

        sum_buy_amt, sum_buy_qty, sum_sell_amt, sum_sell_qty = self.get_monitor_fill_sum_amt()
        buymatch_amt = int(df_group[df_group['來自'] == '盤面交易']['成交金額'].sum())
        buymatch_qty = int(df_group[df_group['來自'] == '盤面交易']['成交張數'].sum())

        if buymatch_amt != sum_buy_amt or buymatch_qty != sum_buy_qty:
            QMessageBox.warning(self, "警告", f"盤面交易金額或張數與CB自營系統不符\n盤面交易: {buymatch_amt:,} 元, {buymatch_qty} 張\nCB自營系統: {sum_buy_amt:,} 元, {sum_buy_qty} 張")
        elif buymatch_amt == sum_buy_amt and buymatch_qty == sum_buy_qty:
            QMessageBox.information(self, "確認", "盤面交易金額或張數與CB自營系統相符")


#====================載入時啟動====================

    def examin_quote_duplicate(self, duplicate_cb): #讀取報價表

        if not duplicate_cb.empty:
            duplicate_cb_list = duplicate_cb['CB代號'].unique().tolist()
            duplicate_msg = f"警告：報價表發現重複的CB代號：{', '.join(duplicate_cb_list)}"
            print(duplicate_msg)
            QMessageBox.warning(None, "報價表重複警告", duplicate_msg)

#====================買進賣出====================

    def show_buy_table(self, df_buy_calculated):
        try:
            # 定義 tday 變數
            tday = datetime.now()
            tdaystr = tday.strftime('%Y%m%d')
            yy = tdaystr[2:4]
            mm = tdaystr[4:6]
            
            # 先檢查必要欄位是否存在
            required_columns = ['客戶ID', 'CB代號', '成交張數', '履約利率%', '成交均價', '權利金百元價', '最終手續費']
            missing_columns = [col for col in required_columns if col not in df_buy_calculated.columns]
            if missing_columns:
                print(f"錯誤：缺少必要欄位: {missing_columns}")
                return pd.DataFrame()
            
            # 轉換為字串格式
            df_buy_calculated = df_buy_calculated.copy()  # 避免修改原始資料
            
            # 設定基本欄位
            df_buy_calculated['錄音日期'] = tdaystr

            if 'SRC' in df_buy_calculated.columns:
                df_buy_calculated['錄音時間'] = np.where(df_buy_calculated['SRC'] == 'E', 'E', '')

            else:
                df_buy_calculated['錄音時間'] = df_buy_calculated['錄音時間']
           
            df_buy_calculated['交易日'] = tdaystr

            # 設定生效日：週五為下週一，其他日期為隔天
            if tday.weekday() == 4:  # 週五
                effective_date = tday + pd.Timedelta(days=3)  # 加3天到下週一
            else:
                effective_date = tday + pd.Timedelta(days=1)  # 加1天到隔天
            df_buy_calculated['生效日'] = effective_date.strftime('%Y%m%d')

            # 設定固定欄位
            df_buy_calculated['提前履約界限日'] = edate(tday, 3).strftime('%Y%m%d')
            df_buy_calculated['提前履約賠償金'] = '0'
            if '波動度' in df_buy_calculated.columns:
                df_buy_calculated['標的波動率'] = df_buy_calculated['波動度']
            else:
                df_buy_calculated['標的波動率'] = '0'
            df_buy_calculated['無風險利率'] = rf
            df_buy_calculated['資金成本'] = '1.425'
            df_buy_calculated['轉債面額'] = '100000'
            df_buy_calculated['選擇權型態'] = 'C'
            df_buy_calculated['選擇權買賣別'] = 'S'
            df_buy_calculated['報價方式'] = '1'
            df_buy_calculated['短契約'] = 'N'
            df_buy_calculated['手續費(營業員)'] = '0'
            df_buy_calculated['手續費(業務單位)'] = df_buy_calculated['最終手續費'].astype(int)
            df_buy_calculated['交易員'] = '10112'
            df_buy_calculated['營業員'] = '10112'
            df_buy_calculated['錄音人員'] = '10112'
            df_buy_calculated['子帳號'] = '751'
            df_buy_calculated['固定端契約編號'] = ' '
            df_buy_calculated['長約附加條款'] = 'N'
            df_buy_calculated['價格事件'] = '70'

            # 定義所需欄位順序
            required_columns_order = [
                '新作契約編號', '上傳序號', '交易類型', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '成交張數', '履約利率%', '成交均價', 
                '權利金百元價', '手續費(業務單位)', '成交金額', '單位權利金', '權利金總額', '錄音日期', '錄音時間', '交易日', '生效日', '交割日期', 
                '選擇權到期日', '賣回日', '賣回價', '提前履約界限日', '提前履約賠償金', '標的波動率', '無風險利率', '資金成本', '轉債面額', 
                '選擇權型態', '選擇權買賣別', '報價方式', '短契約', '手續費(營業員)', '交易員', '營業員', '錄音人員', '子帳號', '固定端契約編號', '長約附加條款', '價格事件', '來自'
            ]
            
            missing_final_columns = [col for col in required_columns_order if col not in df_buy_calculated.columns]
            if missing_final_columns:
                print(f"警告：缺少最終欄位: {missing_final_columns}")
                for col in missing_final_columns:
                    df_buy_calculated[col] = ''  # 設定預設值

            # 重新排列欄位順序
            df_buy_final = df_buy_calculated[required_columns_order]
            
            # 轉換為字串格式
            df_buy_final = df_buy_final.applymap(lambda x: strip_trailing_zeros(x) if pd.notna(x) else '').astype(str)
            
            # 取得現有表格資料並合併
            df_exist = self.get_table_data(self.table_buy)
            
            if not df_exist.empty:
                df_buy_final = pd.concat([df_exist, df_buy_final], ignore_index=True)
            
            conn = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
            df_buy_seq = strip_whitespace(pd.read_sql("SELECT * FROM FSPFLIB.ASPROD ORDER BY PRDID DESC LIMIT 10", conn))
            conn.close()
            
            first_seqno_buy = df_buy_seq.iloc[0]['PRDID']
            last_number_buy = int(str(first_seqno_buy)[-4:])
            df_buy_final['新作契約編號'] = [f"ASOP{yy}{mm}{last_number_buy + i + 1:04d}" for i in range(len(df_buy_final))]
            df_buy_final['上傳序號'] = [f"A{tdaystr}{i+1:03d}" for i in range(len(df_buy_final))]



            self.table_buy.setRowCount(len(df_buy_final))
            blue_columns = ['交易類型', '權利金百元價', '履約利率%', '手續費(業務單位)', '成交張數', '成交均價']
            for i, row in df_buy_final.iterrows():
                for j, col in enumerate(df_buy_final.columns):
                    item = QTableWidgetItem(str(row[col]))
                    if col in blue_columns:
                        item.setBackground(QColor(204, 229, 255))  # 淺藍色
                    self.table_buy.setItem(i, j, item)
            
        except Exception as e:
            print(f"顯示買進表格時發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()

    def renumber_buy_table(self):
        """重新編號買進表格的新作契約編號和上傳序號"""
        try:
            # 讀取現有表格資料
            df_buy = self.get_table_data(self.table_buy)
            
            if df_buy.empty:
                QMessageBox.warning(self, "警告", "買進表格沒有資料！")
                return
            
            # 從數據庫獲取最後一個編號
            conn = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
            df_buy_seq = strip_whitespace(pd.read_sql("SELECT * FROM FSPFLIB.ASPROD ORDER BY PRDID DESC LIMIT 10", conn))
            conn.close()
            
            if df_buy_seq.empty:
                QMessageBox.warning(self, "警告", "無法從數據庫獲取編號！")
                return
            
            first_seqno_buy = df_buy_seq.iloc[0]['PRDID']
            last_number_buy = int(str(first_seqno_buy)[-4:])
            
            # 獲取當前日期
            tday = datetime.now()
            tdaystr = tday.strftime('%Y%m%d')
            yy = tdaystr[2:4]
            mm = tdaystr[4:6]
            
            # 重新生成新作契約編號和上傳序號
            df_buy['新作契約編號'] = [f"ASOP{yy}{mm}{last_number_buy + i + 1:04d}" for i in range(len(df_buy))]
            df_buy['上傳序號'] = [f"A{tdaystr}{i+1:03d}" for i in range(len(df_buy))]
            
            # 更新表格顯示
            buy_columns = [self.table_buy.horizontalHeaderItem(i).text() for i in range(self.table_buy.columnCount())]
            self.table_buy.setRowCount(len(df_buy))
            
            blue_columns = ['交易類型', '權利金百元價', '履約利率%', '手續費(業務單位)', '成交張數', '成交均價']
            for i, row in df_buy.iterrows():
                for j, col in enumerate(df_buy.columns):
                    if col in buy_columns:
                        col_idx = buy_columns.index(col)
                        item = QTableWidgetItem(str(row[col]))
                        if col in blue_columns:
                            item.setBackground(QColor(204, 229, 255))  # 淺藍色
                        self.table_buy.setItem(i, col_idx, item)
            
            QMessageBox.information(self, "成功", f"已重新編號 {len(df_buy)} 筆買進資料！")
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"重新編號時發生錯誤：{e}")
            import traceback
            traceback.print_exc()

    def renumber_sell_table(self):
        """重新編號賣出表格的解約契約編號"""
        try:
            # 讀取現有表格資料
            df_sell = self.get_table_data(self.table_sell)
            
            if df_sell.empty:
                QMessageBox.warning(self, "警告", "賣出表格沒有資料！")
                return
            
            # 從數據庫獲取最後一個編號
            conn = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
            df_sell_seq = strip_whitespace(pd.read_sql("SELECT * FROM FSPFLIB.ASSURR ORDER BY SEQNO DESC LIMIT 10", conn))
            conn.close()
            
            if df_sell_seq.empty:
                QMessageBox.warning(self, "警告", "無法從數據庫獲取編號！")
                return
            
            first_seqno = df_sell_seq.iloc[0]['SEQNO']
            last_number = int(str(first_seqno)[-4:])
            
            # 獲取當前日期
            tday = datetime.now()
            tdaystr = tday.strftime('%Y%m%d')
            yy = tdaystr[2:4]
            mm = tdaystr[4:6]
            
            # 重新生成解約契約編號
            df_sell['解約契約編號'] = [f"ASCP{yy}{mm}{last_number + i + 1:04d}" for i in range(len(df_sell))]
            
            # 更新表格顯示
            sell_columns = [self.table_sell.horizontalHeaderItem(i).text() for i in range(self.table_sell.columnCount())]
            self.table_sell.setRowCount(len(df_sell))
            
            for i, row in df_sell.iterrows():
                for j, col in enumerate(df_sell.columns):
                    if col in sell_columns:
                        col_idx = sell_columns.index(col)
                        item = QTableWidgetItem(str(row[col]))
                        self.table_sell.setItem(i, col_idx, item)
            
            QMessageBox.information(self, "成功", f"已重新編號 {len(df_sell)} 筆賣出資料！")
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"重新編號時發生錯誤：{e}")
            import traceback
            traceback.print_exc()

    def get_prepay_amount(self, df_sell):
        """計算提前履約賠償金（含除錯輸出）"""
        today_dt = pd.to_datetime(datetime.now().strftime('%Y%m%d'), format='%Y%m%d')

        if '提前履約界限日' in df_sell.columns:
            prepay_limit_dt = pd.to_datetime(df_sell['提前履約界限日'].astype(str), format='%Y%m%d', errors='coerce')
            prepay_condition = prepay_limit_dt >= today_dt

        else:
            prepay_condition = pd.Series(False, index=df_sell.index)

        prepay_amount = pd.to_numeric(df_sell.get('提前履約賠償金', 0), errors='coerce').fillna(0).astype(float)
        df_sell['提前履約賠償金'] = np.where(prepay_condition.fillna(False), prepay_amount, 0.0)
        df_sell['履約價'] = pd.to_numeric(df_sell['履約價'], errors='coerce').fillna(0) + df_sell['提前履約賠償金']
        df_sell['提前履約賠償金'] = df_sell['提前履約賠償金'] * df_sell['履約張數'].astype(float).astype(int) * 1000

        cols = []

        if '提前履約界限日' in df_sell.columns:
            cols.append('提前履約界限日')
        cols += ['提前履約賠償金','履約價']

        if '提前履約界限日' in df_sell.columns:
            df_sell.drop(columns=['提前履約界限日'], inplace=True)
        return df_sell
    
    def show_sell_table(self, df_sell_data, from_where=None):
        """整理賣出資訊，並合併現有資料"""
        try:
            conn = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
            df_cusname = strip_whitespace(pd.read_sql("SELECT CUSID, CUSNAME FROM FSPFLIB.FSPCS0M WHERE CBASCODE = 'Y'", conn))
            df_prdid = df_sell_data['原單契約編號'].unique()
            prdid_values = "', '".join(df_prdid)
            
            sql_query = f"""
                    SELECT 
                        CUSID,
                        PRDID,
                        CBTPPRI,
                        CBTPDT,
                        OPTEXDT,
                        PERDATE,
                        PREPAY
                    FROM FSPFLIB.ASPROD 
                    WHERE PRDID IN ('{prdid_values}')
                """
            df_contracts = strip_whitespace(pd.read_sql(sql_query, conn))
            df_contracts.columns = ['客戶ID', '原單契約編號', '賣回價', '賣回日', '選擇權到期日', '提前履約界限日', '提前履約賠償金']
            df_sell_copy = df_sell_data.copy()
            # 合併時保留左表為主，但允許用右表補齊（特別是提前履約賠償金）
            overlap = ['客戶ID','賣回價','賣回日','選擇權到期日','提前履約界限日','提前履約賠償金']
            df_sell_copy = df_sell_copy.merge(
                df_contracts[['原單契約編號'] + overlap],
                on='原單契約編號', how='left', suffixes=('', '_r')
            )
            for c in overlap:
                rcol = f'{c}_r'
                if rcol in df_sell_copy.columns:
                    if c == '提前履約賠償金':
                        left_num = pd.to_numeric(df_sell_copy.get(c), errors='coerce')
                        right_num = pd.to_numeric(df_sell_copy.get(rcol), errors='coerce')
                        use_right = left_num.isna() | (left_num == 0)
                        df_sell_copy[c] = np.where(use_right, right_num, left_num)
                    else:
                        left_val = df_sell_copy.get(c)
                        right_val = df_sell_copy.get(rcol)
                        is_empty = left_val.isna() | (left_val.astype(str).str.strip() == '')
                        df_sell_copy[c] = left_val.mask(is_empty, right_val)
            # 清理右表暫存欄
            drop_cols = [f'{c}_r' for c in overlap if f'{c}_r' in df_sell_copy.columns]
            if drop_cols:
                df_sell_copy.drop(columns=drop_cols, inplace=True)
            # 計算提前履約賠償金與履約價
            df_sell_copy = self.get_prepay_amount(df_sell_copy)
            if '客戶名稱' not in df_sell_copy.columns:
                df_sell_copy = df_sell_copy.merge(df_cusname[['CUSID', 'CUSNAME']], left_on='客戶ID', right_on='CUSID', how='left')
                df_sell_copy.rename(columns={'CUSNAME': '客戶名稱'}, inplace=True)
            if 'CB名稱' not in df_sell_copy.columns:
                df_sell_copy['CB代號'] = df_sell_copy['CB代號'].apply(lambda x: strip_trailing_zeros(x)).astype(str)
                df_sell_copy = df_sell_copy.merge(self.df_cbinfo[['CB代號', 'CB名稱']], left_on='CB代號', right_on='CB代號', how='left')

            df_sell = df_sell_copy[[
                '原單契約編號', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
                '履約張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日', '提前履約賠償金', '履約價', 
                '選擇權交割單價', '交割總金額', '錄音時間', '來自'
            ]]

            if from_where == 'Execution':
                df_sell['交割總金額'] = df_sell['履約張數'] * df_sell['履約價'] * 1000
            else:
                df_sell['選擇權交割單價'] = (pd.to_numeric(df_sell['成交均價'], errors='coerce') - pd.to_numeric(df_sell['履約價'], errors='coerce')).clip(lower=0)
                df_sell['交割總金額'] = df_sell['履約張數'].astype(float).astype(int) * df_sell['選擇權交割單價'].astype(float) * 1000
            
            df_sell['交割總金額'] = np.round(pd.to_numeric(df_sell['交割總金額'], errors='coerce')).astype(int)
            df_sell['選擇權交割單價'] = df_sell['選擇權交割單價'].apply(lambda x: float_to_str_maxlen(x, 11))
            df_sell['成交均價'] = df_sell['成交均價'].apply(lambda x: float_to_str_maxlen(x, 11))

            df_exist = self.get_table_data(self.table_sell)
            df_sell = pd.concat([df_exist, df_sell], ignore_index=True)
            
            df_sell_seq = strip_whitespace(pd.read_sql("SELECT * FROM FSPFLIB.ASSURR ORDER BY SEQNO DESC LIMIT 10", conn)) #取解約契約編號
            first_seqno = df_sell_seq.iloc[0]['SEQNO']
            last_number = int(str(first_seqno)[-4:])
            tday = datetime.now()
            tdaystr = tday.strftime('%Y%m%d')
            yy = tdaystr[2:4]
            mm = tdaystr[4:6]
            new_seqnos_sell = [f"ASCP{yy}{mm}{last_number + i + 1:04d}" for i in range(len(df_sell))]
            #df_sell.insert(0, '解約契約編號', new_seqnos_sell)
            df_sell['解約契約編號'] = new_seqnos_sell
            df_sell_final = df_sell.applymap(lambda x: strip_trailing_zeros(x)).astype(str)
            # 檢查成交均價是否都大於履約價
            try:
                price_check = pd.to_numeric(df_sell_final['成交均價'], errors='coerce') > pd.to_numeric(df_sell_final['履約價'], errors='coerce')
                if not price_check.all():
                    invalid_rows = df_sell_final[~price_check]
                    error_msg = "以下資料的成交均價小於履約價:\n"
                    for _, row in invalid_rows.iterrows():
                        error_msg += f"契約編號: {row['原單契約編號']}, 成交均價: {row['成交均價']}, 履約價: {row['履約價']}\n"
                    QMessageBox.warning(self, "價格檢查警告", error_msg)
            except Exception as e:
                print(f"價格檢查時發生錯誤: {e}")

            self.table_sell.setRowCount(len(df_sell_final))
            blue_columns = ['原單契約編號', '履約張數', '成交均價', '履約利率']
            for i, row in df_sell_final.iterrows():
                for j, col in enumerate(df_sell_final.columns):
                    item = QTableWidgetItem(str(row[col]))
                    if col in blue_columns:
                        item.setBackground(QColor(204, 229, 255))  # 淺藍色
                    self.table_sell.setItem(i, j, item)
            conn.close()
        except Exception as e:
            print(f"顯示賣出表格時發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()

#====================刷新資料====================

    def refresh_data(self):
        try:
            # 取得 UI 選擇的交割日
            self.loading_vips()
            settle_qdate = self.dateedit_settle.date()
            settle_date = datetime(settle_qdate.year(), settle_qdate.month(), settle_qdate.day())
            tday = datetime.now()
            df_buy = read_today_trade_buy(tday, settle_date)
            df_sell = read_today_trade_sell(tday, settle_date)
            df_buy = calculate_new_trade_batch(df_buy)
            self.show_buy_table(df_buy)
            self.show_sell_table(df_sell)
            QMessageBox.information(self, "刷新成功", "資料已重新載入！")
        except Exception as e:
            QMessageBox.critical(self, "刷新失敗", f"發生錯誤：{e}")

    def refresh_quote(self):
        """重新載入報價表"""
        try:
            self.df_quote, self.duplicate_cb, self.df_cbinfo = load_quote()
            QMessageBox.information(self, "重新載入成功", "報價表已重新載入完成！")
            print("報價表已重新載入完成")
        except Exception as e:
            QMessageBox.critical(self, "載入失敗", f"無法載入報價表：{e}")

    def open_quote_file(self):
        """打開報價表Excel文件"""
        try:
            quote_file_path = r'\\10.72.228.112\cbas業務公用區\統一證CBAS報價表_內部.xlsm'
            if os.path.exists(quote_file_path):
                os.startfile(quote_file_path)
            else:
                QMessageBox.warning(self, "警告", f"報價表文件不存在：\n{quote_file_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"打開報價表時發生錯誤：{e}")

    def show_quote_table(self):
        """顯示報價表窗口"""
        try:
            # 如果窗口已經存在且仍然顯示，則將其帶到前台
            if self.quote_window is not None and self.quote_window.isVisible():
                self.quote_window.activateWindow()
                self.quote_window.raise_()
                return
            
            # 檢查是否有報價表資料
            if not hasattr(self, 'df_quote') or self.df_quote.empty:
                QMessageBox.warning(self, "警告", "報價表資料為空！請先重新載入報價表。")
                return
            
            # 創建新的報價表窗口
            self.quote_window = QuoteTableWindow(self.df_quote.copy())
            self.quote_window.show()
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"顯示報價表失敗：{e}")
            print(f"顯示報價表錯誤：{e}")
    
    def show_quote_calculator(self):
        """顯示報價計算機窗口"""
        try:
            # 檢查是否有報價表資料
            if not hasattr(self, 'df_quote') or self.df_quote.empty:
                QMessageBox.warning(self, "警告", "報價表資料為空！請先重新載入報價表。")
                return
            
            # 創建新的報價計算機窗口
            self.calculator_window = QuoteCalculatorWindow(self.df_quote.copy())
            self.calculator_window.show()
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"顯示報價計算機失敗：{e}")
            print(f"顯示報價計算機錯誤：{e}")

#====================議價交易====================

    def get_table_data(self, table):
        """從表格中讀取所有資料並轉換為 DataFrame"""
        data = []
        headers = [table.horizontalHeaderItem(i).text() for i in range(table.columnCount())]
        for row in range(table.rowCount()):
            row_data = []
            for col in range(table.columnCount()):
                item = table.item(row, col)
                row_data.append(item.text() if item else "")
            if any(cell.strip() for cell in row_data):
                data.append(row_data)
        return pd.DataFrame(data, columns=headers) if data else pd.DataFrame(columns=headers)

    def process_bargain(self):
        """處理議價交易資料"""
        try:
            bargain_data = self.get_table_data(self.table_bargain)
            if bargain_data.empty:
                QMessageBox.warning(self, "警告", "議價交易分頁沒有資料！")
                return
            
            processed_bargain = process_bargain_records(bargain_data, self.df_quote)
            
            if not processed_bargain.empty:
                self.update_bargain_table(processed_bargain)
                QMessageBox.information(self, "處理成功", f"已處理 {len(processed_bargain)} 筆議價交易資料！")
            else:
                QMessageBox.warning(self, "處理結果", "沒有可處理的議價交易資料！")
        except Exception as e:
            QMessageBox.critical(self, "處理失敗", f"發生錯誤：{e}")

    def update_bargain_table(self, df_bargain_processed):
        """將處理後的議價交易資料更新回表格"""
        if df_bargain_processed.empty:
            return
        
        self.table_bargain.setRowCount(len(df_bargain_processed))
        
        #light_yellow = QColor(255, 255, 224)
        light_green = QColor(204, 255, 204)
        light_blue = QColor(204, 229, 255)
        for i, row in df_bargain_processed.iterrows():
            for j, col in enumerate(df_bargain_processed.columns):
                item = QTableWidgetItem(str(row[col]))
                if col in ['交割日期', '客戶名稱', 'CB名稱', '議價金額', '銀行', '分行', '銀行帳號', '集保帳號', '通訊地址']:
                    item.setBackground(light_blue)

                if col in ['成交日期', '錄音時間', '單據編號', 'T+?交割', '買/賣', '客戶ID', 'CB代號', '議價張數', '議價價格', '參考價', '錄音時間']:
                    item.setBackground(light_green)

                self.table_bargain.setItem(i, j, item)

    def generate_tickets(self):
        """產給付憑證及買賣成交單"""
        try:
            bargain_data = self.get_table_data(self.table_bargain)
            if bargain_data.empty:
                QMessageBox.warning(self, "警告", "議價交易分頁沒有資料！")
                return
            
            for index, row in bargain_data.iterrows():
                if any(cell.strip() for cell in row if isinstance(cell, str)):
                    row_dict = row.to_dict()
                    generate_settlement_voucher(row_dict)
                    generate_trading_slip(row_dict)

            generate_bargain_upload_file(bargain_data)
            save_trading_statement(bargain_data)
            
            QMessageBox.information(self, "生成成功", f"已生成 {len(bargain_data)} 筆交易的給付憑證及成交單！\n\n已將議價交易資料儲存至議價明細.xlsx")
        except Exception as e:
            QMessageBox.critical(self, "生成失敗", f"發生錯誤：{e}")

    def add_bargain_to_new_trade(self):
        """將議價交易添加到新作買進分頁"""
        try:
            if self.table_bargain.rowCount() == 0:
                QMessageBox.warning(self, "警告", "請先處理議價交易資料！")
                return
            
            bargain_data = self.get_table_data(self.table_bargain)
            if bargain_data.empty:
                QMessageBox.warning(self, "警告", "議價交易表格沒有資料！")
                return
            
            bargain_data_buy = bargain_data[bargain_data['買/賣'] == '買']
            if not bargain_data_buy.empty:
                bargain_data_buy['成交張數'] = pd.to_numeric(bargain_data_buy['議價張數'], errors='coerce').fillna(0)
                bargain_data_buy['成交均價'] = pd.to_numeric(bargain_data_buy['議價價格'], errors='coerce').fillna(0)
                bargain_data_buy['成交金額'] = bargain_data_buy['議價金額'].str.replace(',', '', regex=False)
                bargain_data_buy['交易類型'] = 'ASO'
                convert_to_aso = calculate_new_trade_batch(bargain_data_buy)
                convert_to_aso['來自'] = '議價交易'
                self.show_buy_table(convert_to_aso)
            
            bargain_data_sell = bargain_data[bargain_data['買/賣'] == '賣']
            if not bargain_data_sell.empty:
                bargain_data_sell['成交張數'] = pd.to_numeric(bargain_data_sell['議價張數'], errors='coerce').fillna(0)
                bargain_data_sell['成交均價'] = pd.to_numeric(bargain_data_sell['議價價格'], errors='coerce').fillna(0)
                bargain_data_sell['成交金額'] = bargain_data_sell['議價金額'].str.replace(',', '', regex=False)
                convert_to_sell = bargain_sell(bargain_data_sell)
                convert_to_sell['來自'] = '議價交易'
                self.show_sell_table(convert_to_sell)


            
            QMessageBox.information(self, "成功", f"已添加 {len(bargain_data)} 筆議價交易！")
        except Exception as e:
            QMessageBox.critical(self, "新增失敗", f"發生錯誤：{e}")
 
#====================實物履約====================

    def query_exercise_info(self):
        """查詢履約資訊"""
        try:
            # 獲取輸入值
            cus_id_text = self.input_cus_id.currentText().strip()
            cb_code_text = self.input_cb_code.currentText().strip()
            exercise_qty = self.input_exercise_qty.text().strip()
            # 從ComboBox中提取客戶ID和CB代號
            
            cus_id = cus_id_text.split(" - ")[0] if " - " in cus_id_text else cus_id_text  # 客戶ID取分隔符號前的部分
            cb_code = cb_code_text.split(" - ")[0] if " - " in cb_code_text else cb_code_text  # CB代號取分隔符號前的部分
            
            # 獲取交割日
            settlement_qdate = self.input_settlement_date.date()
            settlement_date = datetime(settlement_qdate.year(), settlement_qdate.month(), settlement_qdate.day())
            
            # 驗證輸入
            if not cus_id:
                QMessageBox.warning(self, "輸入錯誤", "請輸入客戶ID！")
                return
            if not cb_code:
                QMessageBox.warning(self, "輸入錯誤", "請輸入CB代號！")
                return
            if not exercise_qty:
                QMessageBox.warning(self, "輸入錯誤", "請輸入履約張數！")
                return
                
            try:
                exercise_qty_int = int(exercise_qty)
                if exercise_qty_int <= 0:
                    QMessageBox.warning(self, "輸入錯誤", "履約張數必須大於0！")
                    return
            except ValueError:
                QMessageBox.warning(self, "輸入錯誤", "履約張數必須是整數！")
                return
            
            # 查詢資料庫獲取相關契約資訊
            result_data = fetch_exercise_contracts(cus_id, cb_code, exercise_qty_int, settlement_date, self.df_quote)
            
            if result_data.empty:
                QMessageBox.information(self, "查詢結果", "沒有找到符合條件的契約資料！")
                self.table_exercise_result.setRowCount(0)
                return
            
            # 更新結果表格
            update_exercise_result_table(self.table_exercise_result, result_data)
            
            QMessageBox.information(self, "查詢成功", f"找到 {len(result_data)} 筆符合條件的契約！")
            
        except Exception as e:
            QMessageBox.critical(self, "查詢失敗", f"發生錯誤：{e}")
    
#====================選擇權續期====================

    def query_renewal_contracts(self):
        """查詢續期合約資料並進行聚合"""
        cus_id_text = self.input_renewal_cus_id.currentText().strip()
        cb_code_text = self.input_renewal_cb_code.currentText().strip()
        self.df_original_contracts = query_renewal_contracts(cus_id_text, cb_code_text, self.df_quote, self.table_renewal_query)

            
    #def update_renewal_table(self, table, df, columns):
    #    """通用方法：用DataFrame更新表格"""
    #    update_renewal_table(table, df, columns)
    


if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = TableEditor()
    # 不自動讀資料，啟動時不呼叫 get_today_trade_buy/get_today_trade_sell
    editor.show()
    sys.exit(app.exec_())








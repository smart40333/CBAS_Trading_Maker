import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QLineEdit, QPushButton
from format_utils import next_business_day


class QuoteCalculatorWindow(QWidget):
    """報價計算機窗口"""
    def __init__(self, df_quote):
        super().__init__()
        self.df_quote = df_quote
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle("報價計算機")
        self.setGeometry(300, 300, 400, 300)
        
        layout = QVBoxLayout()
        
        # CB標的選擇
        cb_layout = QHBoxLayout()
        cb_layout.addWidget(QLabel("CB標的:"))
        self.cb_combo = QComboBox()
        self.cb_combo.setEditable(True)
        # 添加CB代號和名稱到下拉選單
        for _, row in self.df_quote.iterrows():
            cb_text = f"{row['CB代號']} - {row['CB名稱']}"
            self.cb_combo.addItem(cb_text, row['CB代號'])
        self.cb_combo.currentTextChanged.connect(self.on_cb_changed)
        cb_layout.addWidget(self.cb_combo)
        layout.addLayout(cb_layout)
        
        # CB到期日
        due_layout = QHBoxLayout()
        due_layout.addWidget(QLabel("CB賣回日:"))
        self.due_date_edit = QLineEdit()
        self.due_date_edit.setReadOnly(True)
        due_layout.addWidget(self.due_date_edit)
        layout.addLayout(due_layout)
        
        # 履約利率
        rate_layout = QHBoxLayout()
        rate_layout.addWidget(QLabel("履約利率:"))
        self.rate_edit = QLineEdit()
        rate_layout.addWidget(self.rate_edit)
        layout.addLayout(rate_layout)
        
        # 手續費
        fee_layout = QHBoxLayout()
        fee_layout.addWidget(QLabel("手續費:"))
        self.fee_edit = QLineEdit("150")
        fee_layout.addWidget(self.fee_edit)
        layout.addLayout(fee_layout)

        # 張數
        qty_layout = QHBoxLayout()
        qty_layout.addWidget(QLabel("張數:"))
        self.qty_edit = QLineEdit("1")
        qty_layout.addWidget(self.qty_edit)
        layout.addLayout(qty_layout)

        # 成交均價
        price_layout = QHBoxLayout()
        price_layout.addWidget(QLabel("成交均價:"))
        self.price_edit = QLineEdit("100")
        price_layout.addWidget(self.price_edit)
        layout.addLayout(price_layout)
        
        # 計算按鈕
        calc_btn = QPushButton("計算")
        calc_btn.clicked.connect(self.calculate)
        layout.addWidget(calc_btn)
        
        # 計算結果區域
        # 百元報價結果
        hundred_layout = QHBoxLayout()
        hundred_layout.addWidget(QLabel("百元報價:"))
        self.hundred_label = QLabel("")
        self.hundred_label.setStyleSheet("font-size: 14px; font-weight: bold; color: blue;")
        hundred_layout.addWidget(self.hundred_label)
        hundred_layout.addStretch()
        layout.addLayout(hundred_layout)
        
        # 成交金額結果
        price_layout = QHBoxLayout()
        price_layout.addWidget(QLabel("成交金額:"))
        self.price_label = QLabel("")
        self.price_label.setStyleSheet("font-size: 14px; font-weight: bold; color: green;")
        price_layout.addWidget(self.price_label)
        price_layout.addStretch()
        layout.addLayout(price_layout)
        
        # 單位權利金結果
        premium_layout = QHBoxLayout()
        premium_layout.addWidget(QLabel("單位權利金:"))
        self.premium_label = QLabel("")
        self.premium_label.setStyleSheet("font-size: 14px; font-weight: bold; color: purple;")
        premium_layout.addWidget(self.premium_label)
        premium_layout.addStretch()
        layout.addLayout(premium_layout)
        
        # 權利金總額結果
        total_layout = QHBoxLayout()
        total_layout.addWidget(QLabel("權利金總額:"))
        self.total_label = QLabel("")
        self.total_label.setStyleSheet("font-size: 14px; font-weight: bold; color: red;")
        total_layout.addWidget(self.total_label)
        total_layout.addStretch()
        layout.addLayout(total_layout)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def on_cb_changed(self, text):
        """當CB標的改變時，更新相關欄位"""
        try:
            # 從選項中提取CB代號
            cb_code = text.split(" - ")[0] if " - " in text else text
            
            # 在報價表中查找對應的資料
            matching_rows = self.df_quote[self.df_quote['CB代號'] == cb_code]
            if not matching_rows.empty:
                row = matching_rows.iloc[0]
                self.due_date_edit.setText(str(row['賣回日']))
                self.rate_edit.setText(str(row['履約利率']))
        except Exception as e:
            print(f"更新CB資料時發生錯誤: {e}")
    
    def calculate(self):
        """計算百元報價"""
        try:
            # 獲取輸入值
            cb_code = self.cb_combo.currentText().split(" - ")[0] if " - " in self.cb_combo.currentText() else self.cb_combo.currentText()
            due_date_str = self.due_date_edit.text()
            rate = float(self.rate_edit.text())
            fee = float(self.fee_edit.text())
            qty = float(self.qty_edit.text())
            price = float(self.price_edit.text())
            
            # 計算年期
            due_date = pd.to_datetime(due_date_str, format='%Y%m%d')
            t_plus_2 = pd.Timestamp(next_business_day(datetime.now(), 2).date())
            year_period = ((due_date - t_plus_2).days + 1) / 365

            # 獲取賣回價
            matching_rows = self.df_quote[self.df_quote['CB代號'] == cb_code]
            if not matching_rows.empty:
                sellback_price = float(matching_rows.iloc[0]['賣回價'])
            else:
                sellback_price = 100  # 預設值
            
            # 計算百元報價
            hundred_price = round(rate * year_period - (sellback_price - 100), 2) + fee / 1000
            price_total = qty * price * 1000
            premium_per = price - 100 + hundred_price
            premium_total = premium_per * qty * 1000

            # 更新各個結果標籤
            self.hundred_label.setText(f"{hundred_price:.2f}")
            self.price_label.setText(f"{price_total:,.0f}")
            self.premium_label.setText(f"{premium_per:.2f}")
            self.total_label.setText(f"{premium_total:,.0f}")

            
        except Exception as e:
            # 清空所有結果標籤並顯示錯誤訊息
            self.hundred_label.setText("錯誤")
            self.price_label.setText("錯誤")
            self.premium_label.setText("錯誤")
            self.total_label.setText("錯誤")
            print(f"計算錯誤: {e}") 
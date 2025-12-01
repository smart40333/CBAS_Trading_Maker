import pandas as pd
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtCore import Qt


class QuoteTableWindow(QWidget):
    """報價表顯示窗口"""
    def __init__(self, df_quote):
        super().__init__()
        self.df_quote = df_quote
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("報價表查看")
        self.setGeometry(100, 100, 1200, 800)
        
        layout = QVBoxLayout()
        
        # 標題
        title_label = QLabel("CB報價表")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 搜尋區域
        search_frame = QWidget()
        search_layout = QHBoxLayout()
        
        search_layout.addWidget(QLabel("CB代號搜尋："))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("輸入CB代號進行搜尋...")
        self.search_input.textChanged.connect(self.filter_table)
        search_layout.addWidget(self.search_input)
        
        # 重新整理按鈕
        self.refresh_btn = QPushButton("重新整理")
        self.refresh_btn.clicked.connect(self.refresh_data)
        search_layout.addWidget(self.refresh_btn)
        
        search_layout.addStretch()
        search_frame.setLayout(search_layout)
        layout.addWidget(search_frame)
        
        # 報價表格
        self.table = QTableWidget()
        self.setup_table()
        layout.addWidget(self.table)
        
        # 狀態欄
        self.status_label = QLabel(f"共 {len(self.df_quote)} 筆資料")
        layout.addWidget(self.status_label)
        
        self.setLayout(layout)
        
    def setup_table(self):
        """設置表格"""
        if self.df_quote.empty:
            return
            
        self.table.setRowCount(len(self.df_quote))
        self.table.setColumnCount(len(self.df_quote.columns))
        self.table.setHorizontalHeaderLabels(self.df_quote.columns.tolist())
        
        # 填入資料
        for i, (index, row) in enumerate(self.df_quote.iterrows()):
            for j, col in enumerate(self.df_quote.columns):
                value = str(row[col]) if pd.notna(row[col]) else ""
                item = QTableWidgetItem(value)
                # 設置關鍵欄位的背景色
                if col in ['CB代號', 'CB名稱']:
                    item.setBackground(QColor(230, 255, 230))  # 淺綠色
                elif col in ['履約利率', '低履約利率']:
                    item.setBackground(QColor(255, 255, 224))  # 淺黃色
                elif col in ['賣回價', '賣回日']:
                    item.setBackground(QColor(204, 229, 255))  # 淺藍色
                self.table.setItem(i, j, item)
        
        # 調整欄位寬度
        self.table.resizeColumnsToContents()
        
    def filter_table(self):
        """根據CB代號篩選表格"""
        search_text = self.search_input.text().strip().upper()
        
        if not search_text:
            # 顯示所有資料
            filtered_df = self.df_quote
        else:
            # 篩選CB代號包含搜尋文字的資料
            filtered_df = self.df_quote[
                self.df_quote['CB代號'].astype(str).str.upper().str.contains(search_text, na=False) |
                self.df_quote['CB名稱'].astype(str).str.upper().str.contains(search_text, na=False)
            ]
        
        # 更新表格
        self.table.setRowCount(len(filtered_df))
        for i, (index, row) in enumerate(filtered_df.iterrows()):
            for j, col in enumerate(filtered_df.columns):
                value = str(row[col]) if pd.notna(row[col]) else ""
                item = QTableWidgetItem(value)
                # 設置關鍵欄位的背景色
                if col in ['CB代號', 'CB名稱']:
                    item.setBackground(QColor(230, 255, 230))  # 淺綠色
                elif col in ['履約利率', '低履約利率']:
                    item.setBackground(QColor(255, 255, 224))  # 淺黃色
                elif col in ['賣回價', '賣回日']:
                    item.setBackground(QColor(204, 229, 255))  # 淺藍色
                self.table.setItem(i, j, item)
        
        self.status_label.setText(f"顯示 {len(filtered_df)} / {len(self.df_quote)} 筆資料")
        
    def refresh_data(self):
        """重新整理報價表資料"""
        try:
            # 這裡需要重新載入報價表資料
            QMessageBox.information(self, "提示", "請在主程式中重新載入報價表後再開啟此視窗")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"重新整理失敗：{e}") 
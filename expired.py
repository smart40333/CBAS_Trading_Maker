"""
到期相關功能模組
處理合約到期的查詢、格式化和UI操作
"""

import pandas as pd
import numpy as np
from datetime import datetime
from PyQt5.QtWidgets import QMessageBox, QTableWidgetItem

from db_access import get_expired_contracts_db
from file_reader import read_expired_trade_data
from format_utils import format_expired_contract_data


def query_expired_contracts(dateedit_expired, table_expired, df_quote, table_sell=None, get_table_data_func=None):
    """查詢選定日期到期合約並顯示在表格中"""
    try:
        # 從日期選擇器獲取選定的日期
        selected_qdate = dateedit_expired.date()
        selected_date = datetime(selected_qdate.year(), selected_qdate.month(), selected_qdate.day())
        
        df_expired = get_expired_contracts(selected_date, df_quote, table_sell=table_sell, get_table_data_func=get_table_data_func)
        
        if df_expired.empty:
            date_str = selected_date.strftime('%Y/%m/%d')
            QMessageBox.information(None, "查詢結果", f"選定日期（{date_str}）沒有到期的未賣完合約")
            table_expired.setRowCount(0)
            return
        
        # 更新表格
        table_expired.setRowCount(len(df_expired))

        for i, (index, row) in enumerate(df_expired.iterrows()):
            for j, col in enumerate(df_expired.columns):
                cell = row[col]
                # 若值為 Series/陣列（例如重複欄位），取第一個元素為顯示值
                if isinstance(cell, (pd.Series, np.ndarray, list, tuple)):
                    cell = cell[0] if len(cell) > 0 else ""
                value = "" if pd.isna(cell) else str(cell)
                item = QTableWidgetItem(value)
                table_expired.setItem(i, j, item)
        
        QMessageBox.information(None, "查詢完成", f"找到 {len(df_expired)} 個到期未賣完的合約")
        
    except Exception as e:
        QMessageBox.critical(None, "查詢失敗", f"發生錯誤：{e}")
        print(f"查詢到期合約錯誤：{e}")



def get_expired_contracts(target_date=None, df_quote=None, table_sell=None, get_table_data_func=None):
    """取得已到期且未完全賣出的契約（用於表格顯示）"""
    if target_date is None:
        tday = datetime.now().strftime('%Y%m%d')
    else:
        if isinstance(target_date, str):
            tday = target_date
        else:
            tday = target_date.strftime('%Y%m%d')
    
    # 從資料庫取得到期契約
    df_expired = get_expired_contracts_db(tday) #欄位為400名
    df_expired.rename(columns={'PRDID': '原單契約編號'}, inplace=True)
    df_expired['STORQTY'] = pd.to_numeric(df_expired['STORQTY'], errors='coerce').fillna(0).astype(int)

    
    # 檢查是否有到期的合約
    if df_expired.empty:
        print(f"今日（{tday}）沒有到期的合約")
        return pd.DataFrame(columns=[
            '原單契約編號', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
            '原庫存張數', '今日賣出張數', '剩餘到期張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日',
            '提前履約賠償金', '履約價', '選擇權交割單價', '交割總金額', '錄音時間'
        ])
    
    # 讀取今天的交易資料（中台檔
    df_sell_from_ASCCSV02 = read_expired_trade_data() #欄位已變更為中文
    
    # 讀取提解賣出表格資料（若有傳入）
    if table_sell is not None and get_table_data_func is not None:
        try:
            df_sell_from_table = get_table_data_func(table_sell)
        except Exception:
            df_sell_from_table = pd.DataFrame()
    else:
        df_sell_from_table = pd.DataFrame()
    

    df_result = df_expired.merge(df_quote[['CB代號', 'CB名稱']], left_on='CBCODE', right_on='CB代號', how='left') #找CB名稱
    if not df_sell_from_ASCCSV02.empty:
        df_result = df_result.merge(df_sell_from_ASCCSV02, on='原單契約編號', how='left')

    
    # 若有表格資料，將「非盤面交易」的數量納入扣除
    if not df_sell_from_table.empty:
        # 僅統計 來自 != '盤面交易' 的筆數（如議價交易等），彙總每個原單契約的履約張數
        df_tbl = df_sell_from_table.copy()
        if '來自' in df_tbl.columns:
            df_tbl = df_tbl[df_tbl['來自'] != '盤面交易']
            df_tbl['履約張數'] = pd.to_numeric(df_tbl['履約張數'], errors='coerce').fillna(0)
            df_non_market_qty = (
                df_tbl.groupby('原單契約編號')['履約張數']
                .sum()
                .reset_index()
                .rename(columns={'履約張數': '非盤面賣出張數'})
            )
            df_result = df_result.merge(df_non_market_qty, on='原單契約編號', how='left')
        else:
            df_result['非盤面賣出張數'] = 0
    else:
        df_result['非盤面賣出張數'] = 0

    # 確保基礎欄位存在且為Series
    if '今日賣出張數_ASCCSV02' not in df_result.columns:
        df_result['今日賣出張數_ASCCSV02'] = 0
    if '非盤面賣出張數' not in df_result.columns:
        df_result['非盤面賣出張數'] = 0

    # 以中台檔的今日賣出張數為基礎，再加上非盤面賣出張數
    df_result['今日賣出張數_ASCCSV02'] = pd.to_numeric(df_result['今日賣出張數_ASCCSV02'], errors='coerce').fillna(0).astype(int)
    df_result['非盤面賣出張數'] = pd.to_numeric(df_result['非盤面賣出張數'], errors='coerce').fillna(0).astype(int)
    base_sold = pd.to_numeric(df_result['今日賣出張數_ASCCSV02'], errors='coerce').fillna(0)
    extra_sold = pd.to_numeric(df_result['非盤面賣出張數'], errors='coerce').fillna(0)
    df_result['今日賣出張數'] = base_sold + extra_sold
    df_result['STORQTY'] = pd.to_numeric(df_result.get('STORQTY'), errors='coerce').fillna(0)
    df_result['剩餘到期張數'] = df_result['STORQTY'] - df_result['今日賣出張數']
    df_unsold_contracts = df_result[df_result['剩餘到期張數'] > 0].copy()


    # 使用format_utils中的函數格式化資料
    df_final = format_expired_contract_data(df_unsold_contracts, tday, df_quote)
    
    # 確保剩餘到期張數欄位存在
    if '剩餘到期張數' not in df_final.columns:
        df_final['剩餘到期張數'] = df_final['原庫存張數'] - df_final.get('今日賣出張數', 0)
    
    # 選擇最終輸出欄位
    final_columns = [
        '原單契約編號', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
        '原庫存張數', '今日賣出張數', '剩餘到期張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日',
        '提前履約賠償金', '履約價', '選擇權交割單價', '交割總金額', '錄音時間'
    ]

    # 確保所有欄位都存在
    for col in final_columns:
        if col not in df_final.columns:
            df_final[col] = ''
    
    df_final = df_final[final_columns]

    return df_final


def add_expired_to_sell(table_expired, get_table_data_func, show_sell_table_func):
    """將合約到期添加到提解賣出分頁"""
    try:
        # 檢查合約到期表格是否有資料
        if table_expired.rowCount() == 0:
            QMessageBox.warning(None, "警告", "請先查詢今日到期合約！")
            return
        
        # 取得合約到期表格的資料
        expired_data = get_table_data_func(table_expired)
        expired_data['今日賣出張數'] = pd.to_numeric(expired_data['今日賣出張數'], errors='coerce').fillna(0).astype(int)
        expired_data['剩餘到期張數'] = pd.to_numeric(expired_data['剩餘到期張數'], errors='coerce').fillna(0).astype(int)
        today_sell_data = expired_data['今日賣出張數'].sum()
        today_expired_data = expired_data['剩餘到期張數'].sum()
        reply = QMessageBox.question(None, "資料確認", f"今日賣出: {today_sell_data}張\n今日到期: {today_expired_data}張\n請確認是否正確。", 
                                    QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.No:
            return

        if expired_data.empty:
            QMessageBox.warning(None, "警告", "合約到期表格沒有資料！")
            return
        
        # 轉換為提解賣出格式 - 合約到期的資料已經包含大部分需要的欄位
        sell_columns = [
            '原單契約編號', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
            '履約張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日', '提前履約賠償金', '履約價', 
            '選擇權交割單價', '交割總金額', '錄音時間'
        ]

        # 複製到期資料並進行欄位對應
        expired_data_copy = expired_data.copy()
        
        # 將剩餘到期張數對應到履約張數
        expired_data_copy['履約張數'] = expired_data_copy['剩餘到期張數']
        
        # 確保所有必要欄位都存在
        for col in sell_columns:
            if col not in expired_data_copy.columns:
                expired_data_copy[col] = ''
        
        # 重新排列欄位順序
        expired_for_sell = expired_data_copy[sell_columns]
        expired_for_sell['來自'] = '到期'
        
        # 調用show_sell_table進行處理
        show_sell_table_func(expired_for_sell, from_where='Expired')
        QMessageBox.information(None, "成功", f"已將 {len(expired_data)} 筆到期合約添加到提解賣出分頁！")
        
    except Exception as e:
        QMessageBox.critical(None, "新增失敗", f"發生錯誤：{e}")
        print(f"新增到期合約至提解錯誤：{e}") 
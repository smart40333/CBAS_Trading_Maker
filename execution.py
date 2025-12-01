import pandas as pd
import pyodbc
from datetime import datetime
from PyQt5.QtWidgets import QMessageBox, QTableWidgetItem
from PyQt5.QtGui import QColor
from db_access import strip_whitespace
from format_utils import strip_trailing_zeros


def setup_exercise_input_search(input_cus_id, input_cb_code, df_quote, get_customer_list_func):
    """設置實物履約輸入框的搜尋功能"""
    # 設置客戶ID搜尋功能
    customer_list = get_customer_list_func()
    for customer in customer_list:
        input_cus_id.addItem(customer)
    
    # 設置CB代號搜尋功能
    if df_quote is not None:
        for _, row in df_quote.iterrows():
            cb_code = str(row.get('CB代號', '')).strip()
            cb_name = str(row.get('CB名稱', '')).strip()
            if cb_code and cb_name:
                input_cb_code.addItem(f"{cb_code} - {cb_name}")


def filter_customer_items(input_cus_id, search_text, get_customer_list_func):
    """過濾客戶ID項目"""
    # 如果正在更新中，跳過
    if hasattr(input_cus_id, '_updating_customer') and input_cus_id._updating_customer:
        return
    
    input_cus_id._updating_customer = True
    
    try:
        # 保存當前選擇的項目
        current_text = input_cus_id.currentText()
        
        if not search_text:
            # 如果搜尋文字為空，顯示所有項目
            input_cus_id.clear()
            customer_list = get_customer_list_func()
            for customer in customer_list:
                input_cus_id.addItem(customer)
        else:
            # 清空並重新載入符合條件的項目
            input_cus_id.clear()
            search_text = search_text.strip().upper()
            
            customer_list = get_customer_list_func()
            for customer in customer_list:
                if search_text in customer.upper():
                    input_cus_id.addItem(customer)
        
        # 嘗試恢復之前的選擇
        if current_text and input_cus_id.findText(current_text) >= 0:
            input_cus_id.setCurrentText(current_text)
            
    except Exception as e:
        print(f"過濾客戶ID時發生錯誤：{e}")
    finally:
        input_cus_id._updating_customer = False


def filter_cb_items(input_cb_code, search_text, df_quote):
    """過濾CB代號項目"""
    # 如果正在更新中，跳過
    if hasattr(input_cb_code, '_updating_cb') and input_cb_code._updating_cb:
        return
    
    input_cb_code._updating_cb = True
    
    try:
        # 保存當前選擇的項目
        current_text = input_cb_code.currentText()
        
        if not search_text:
            # 如果搜尋文字為空，顯示所有項目
            input_cb_code.clear()
            if df_quote is not None:
                for _, row in df_quote.iterrows():
                    cb_code = str(row.get('CB代號', '')).strip()
                    cb_name = str(row.get('CB名稱', '')).strip()
                    if cb_code and cb_name:
                        input_cb_code.addItem(f"{cb_code} - {cb_name}")
        else:
            # 清空並重新載入符合條件的項目
            input_cb_code.clear()
            search_text = search_text.strip().upper()
            
            if df_quote is not None:
                for _, row in df_quote.iterrows():
                    cb_code = str(row.get('CB代號', '')).strip()
                    cb_name = str(row.get('CB名稱', '')).strip()
                    
                    # 搜尋CB代號或CB名稱
                    if (search_text in cb_code.upper() or 
                        search_text in cb_name.upper()):
                        input_cb_code.addItem(f"{cb_code} - {cb_name}")
        
        # 嘗試恢復之前的選擇
        if current_text and input_cb_code.findText(current_text) >= 0:
            input_cb_code.setCurrentText(current_text)
            
    except Exception as e:
        print(f"過濾CB代號時發生錯誤：{e}")
    finally:
        input_cb_code._updating_cb = False


def query_exercise_info(input_cus_id, input_cb_code, input_exercise_qty, input_settlement_date, 
                       df_quote, table_exercise_result, fetch_exercise_contracts_func):
    """查詢履約資訊"""
    try:
        # 獲取輸入值
        cus_id_text = input_cus_id.currentText().strip()
        cb_code_text = input_cb_code.currentText().strip()
        exercise_qty = input_exercise_qty.text().strip()
        
        # 從ComboBox中提取客戶ID和CB代號
        cus_id = cus_id_text.split(" - ")[0] if " - " in cus_id_text else cus_id_text
        cb_code = cb_code_text.split(" - ")[0] if " - " in cb_code_text else cb_code_text
        
        # 獲取交割日
        settlement_qdate = input_settlement_date.date()
        settlement_date = datetime(settlement_qdate.year(), settlement_qdate.month(), settlement_qdate.day())
        
        # 驗證輸入
        if not cus_id:
            QMessageBox.warning(None, "輸入錯誤", "請輸入客戶ID！")
            return
        if not cb_code:
            QMessageBox.warning(None, "輸入錯誤", "請輸入CB代號！")
            return
        if not exercise_qty:
            QMessageBox.warning(None, "輸入錯誤", "請輸入履約張數！")
            return
            
        try:
            exercise_qty_int = int(exercise_qty)
            if exercise_qty_int <= 0:
                QMessageBox.warning(None, "輸入錯誤", "履約張數必須大於0！")
                return
        except ValueError:
            QMessageBox.warning(None, "輸入錯誤", "履約張數必須是整數！")
            return
        
        # 查詢資料庫獲取相關契約資訊
        result_data = fetch_exercise_contracts_func(cus_id, cb_code, exercise_qty_int, settlement_date, df_quote)
        
        if result_data.empty:
            QMessageBox.information(None, "查詢結果", "沒有找到符合條件的契約資料！")
            table_exercise_result.setRowCount(0)
            return
        
        # 更新結果表格
        update_exercise_result_table(table_exercise_result, result_data)
        
        QMessageBox.information(None, "查詢成功", f"找到 {len(result_data)} 筆符合條件的契約！")
        
    except Exception as e:
        QMessageBox.critical(None, "查詢失敗", f"發生錯誤：{e}")


def fetch_exercise_contracts(cus_id: str, cb_code: str, exercise_qty: int, settlement_date: datetime, df_quote: pd.DataFrame) -> pd.DataFrame:
    """從資料庫獲取履約契約資訊並進行智能排序分配"""
    try:
        conn = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
        
        # 補齊客戶ID到12位
        db_cusid_length = 12
        cus_id_padded = cus_id.ljust(db_cusid_length)
        
        # 查詢該客戶的現有契約
        sql_query = f"""
            SELECT 
                CUSID,
                PRDID,
                CBCODE,
                STORQTY,
                TRDATE,
                PERRATE,
                CBTPPRI,
                CBTPDT
            FROM FSPFLIB.ASPROD 
            WHERE CUSID = '{cus_id_padded}' 
            AND CBCODE = '{cb_code}' 
            AND STORQTY > 0
        """
        
        df_contracts = strip_whitespace(pd.read_sql(sql_query, conn))
        df_contracts.columns = ['客戶ID', '原單契約編號', 'CB代號', '庫存張數', '成交日期', '原利率', '賣回價', '賣回日']
        
        if df_contracts.empty:
            conn.close()
            return pd.DataFrame()
        
        # 獲取客戶名稱
        df_cusname = strip_whitespace(pd.read_sql(f"SELECT CUSID, CUSNAME FROM FSPFLIB.FSPCS0M WHERE CUSID = '{cus_id_padded}'", conn))
        
        # 獲取CB名稱和相關資訊
        cb_info = df_quote[df_quote['CB代號'] == cb_code]

        conn.close()
        
        # 合併資料
        if not df_cusname.empty:
            df_contracts = pd.merge(df_contracts, df_cusname[['CUSID', 'CUSNAME']], left_on='客戶ID', right_on='CUSID', how='left')
            df_contracts['客戶名稱'] = df_contracts['CUSNAME']
        else:
            df_contracts['客戶名稱'] = ''
        
        if not cb_info.empty:
            df_contracts['CB名稱'] = cb_info.iloc[0]['CB名稱']
        else:
            df_contracts['CB名稱'] = ''
        
        # 數據類型轉換
        df_contracts['原利率'] = pd.to_numeric(df_contracts['原利率'], errors='coerce').fillna(0)
        df_contracts['庫存張數'] = pd.to_numeric(df_contracts['庫存張數'], errors='coerce').fillna(0).astype(int)
        df_contracts['賣回價'] = pd.to_numeric(df_contracts['賣回價'], errors='coerce').fillna(100)
        
        # 轉換日期格式
        df_contracts['成交日期'] = pd.to_datetime(df_contracts['成交日期'], format='%Y%m%d', errors='coerce')
        df_contracts['賣回日'] = pd.to_datetime(df_contracts['賣回日'], format='%Y%m%d', errors='coerce')
        
        # 智能排序：原利率(高→低)、成交日期(早→晚)、庫存張數(少→多)
        df_contracts_sorted = df_contracts.sort_values(
            by=['原利率', '成交日期', '庫存張數'], 
            ascending=[False, True, True]
        ).reset_index(drop=True)
        
        # 創建 df_exe 進行履約分配
        df_exe = create_exercise_allocation(df_contracts_sorted, exercise_qty, settlement_date)
        
        # 檢查是否有庫存不足的情況
        total_exercised = df_exe['此契約履約張數'].sum() if not df_exe.empty else 0
        shortage = exercise_qty - total_exercised
        if shortage > 0:
            QMessageBox.warning(None, "注意", f"客戶可履約張數不足，缺少 {shortage} 張")
        
        # 轉換為顯示格式
        df_exe['客戶ID'] = cus_id
        df_exe['CB代號'] = cb_code
        df_exe['履約張數'] = df_exe['此契約履約張數']
        df_exe['交易日期'] = datetime.now().strftime('%Y%m%d')
        df_exe['交割日期'] = settlement_date.strftime('%Y%m%d')
        df_exe['賣出金額'] = (df_exe['履約價'] * df_exe['此契約履約張數'] * 1000).astype(int)
        df_exe['履約後剩餘張數'] = df_exe['原庫存張數'] - df_exe['此契約履約張數']
        df_exe = df_exe[['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '原單契約編號', '原利率', '交易日期', '交割日期', '成交日期', '履約價', '賣出金額', '原庫存張數', '今日賣出張數', '履約張數', '履約後剩餘張數', '解約類別', '履約方式']]

        return df_exe
        
    except Exception as e:
        print(f"查詢履約契約時發生錯誤: {e}")
        if 'conn' in locals():
            conn.close()
        return pd.DataFrame()


def create_exercise_allocation(df_contracts_sorted, exercise_qty, settlement_date):
    """創建 df_exe 進行履約分配"""
    df_exe = df_contracts_sorted.copy()
    try:
        tday = datetime.now().strftime('%Y%m%d')
        df_had_sold = pd.read_excel(rf"\\10.72.228.83\CBAS_Logs\{tday}\AfterClose\ASCCSV02.csv")
        df_exe = pd.merge(df_exe, df_had_sold, left_on='原單契約編號', right_on='PRDID', how='left')
        df_exe['今日賣出張數'] = df_exe['DEUQTY']
    except Exception as e:
        print(f"讀取已賣出資料時發生錯誤: {e}")
        df_exe['今日賣出張數'] = 0

    # 添加履約相關欄位
    df_exe['此契約履約張數'] = 0
    df_exe['原庫存張數'] = df_exe['庫存張數'].copy()
    df_exe['剩餘張數'] = df_exe['庫存張數'].copy() - df_exe['今日賣出張數']
    df_exe['履約價'] = 0.0
    df_exe['年期'] = 0.0
    
    remaining_qty = exercise_qty
    today = datetime.now().strftime('%Y%m%d')
    
    for idx, row in df_exe.iterrows():
        if remaining_qty <= 0:
            break
        
        available_qty = int(row['庫存張數'])
        exercise_this_contract = min(remaining_qty, available_qty)
        
        # 更新履約張數
        df_exe.at[idx, '此契約履約張數'] = exercise_this_contract
        remaining_contract_qty = available_qty - exercise_this_contract
        df_exe.at[idx, '剩餘張數'] = remaining_contract_qty
        df_exe.at[idx, '解約類別'] = '2' if remaining_contract_qty == 0 else '1'
        df_exe.at[idx, '履約方式'] = '2'
        
        # 計算年期 = (賣回日 - 交割日) / 365
        if pd.notna(row['賣回日']):
            years = ((row['賣回日'] - settlement_date).days + 1)/ 365
            df_exe.at[idx, '年期'] = max(years, 0)  # 確保年期不為負
        else:
            df_exe.at[idx, '年期'] = 0
        
        # 計算履約價 = round(賣回價 - (100 * 年期 * 原利率), 2)
        exercise_price = round(
            row['賣回價'] - (100 * df_exe.at[idx, '年期'] * row['原利率']/100), 2
        )
        df_exe.at[idx, '履約價'] = exercise_price
        
        remaining_qty -= exercise_this_contract
        
        print(f"契約{idx+1}: 原利率={row['原利率']:.4f}, 可用={available_qty}張, 履約={exercise_this_contract}張, 履約價={exercise_price:.2f}")
    
    if remaining_qty > 0:
        print(f"⚠️  客戶可履約張數不足，缺少 {remaining_qty} 張")
    else:
        print(f"✅ 履約分配完成！")
    
    # 只返回有履約的契約
    df_exe_filtered = df_exe[df_exe['此契約履約張數'] > 0].reset_index(drop=True)

    return df_exe_filtered


def update_exercise_result_table(table_exercise_result, df_result):
    """更新履約結果表格"""
    if df_result.empty:
        table_exercise_result.setRowCount(0)
        return
    
    # 設定表格行數
    table_exercise_result.setRowCount(len(df_result))
    
    # 填入資料
    for i, row in df_result.iterrows():
        for j, col in enumerate(df_result.columns):
            item = QTableWidgetItem(str(row[col]))
            
            # 設定表格可編輯
            item.setFlags(item.flags() | 16)  # Qt.ItemIsEditable
            
            # 可以為重要欄位設定背景色
            if col in ['履約張數', '履約價', '賣出金額', '履約後剩餘張數']:
                item.setBackground(QColor(255, 255, 224))  # 淺黃色
            elif col in ['原利率', '原庫存張數']:
                item.setBackground(QColor(230, 255, 230))  # 淺綠色
            
            table_exercise_result.setItem(i, j, item)
    
    print(f"履約結果表格已更新，共 {len(df_result)} 筆資料")


def add_exercise_info(table_exercise_result, table_cbas_to_cb, get_table_data_func):
    """將上方查詢結果添加到下方本日新增履約表格"""
    try:
        # 檢查上方表格是否有資料
        if table_exercise_result.rowCount() == 0:
            QMessageBox.warning(None, "警告", "請先查詢履約資訊！")
            return
        
        # 取得上方表格的當前資料（可能已被用戶修改）
        upper_data = get_table_data_func(table_exercise_result)
        
        if upper_data.empty:
            QMessageBox.warning(None, "警告", "上方表格沒有資料！")
            return
        
        # 取得下方表格現有資料
        current_data = get_table_data_func(table_cbas_to_cb)
        
        # 合併資料
        if current_data.empty:
            combined_data = upper_data
        else:
            combined_data = pd.concat([current_data, upper_data], ignore_index=True)
    
        # ============================備註內容=================================
        combined_data['交割日期'] = pd.to_datetime(combined_data['交割日期'], errors='coerce', format='%Y%m%d')
        combined_data['交易日期'] = pd.to_datetime(combined_data['交易日期'], errors='coerce', format='%Y%m%d')
        combined_data['成交日期'] = pd.to_datetime(combined_data['成交日期'], errors='coerce')

        # 計算天數差
        combined_data['Days'] = (combined_data['交割日期'] - combined_data['交易日期']).dt.days

        # 動態顯示 T+/T-/T0
        def format_T(row):
            days = row['Days']
            if days > 0:
                t_str = f"T+{days}"
            elif days < 0:
                t_str = f"T{days}"  # days 已含負號
            else:
                t_str = "T"
            return f"實物履約{t_str}，{row['客戶名稱']} {row['CB代號']}{row['CB名稱']}，{row['履約張數']}張"

        # 套用
        combined_data['備註'] = combined_data.apply(format_T, axis=1)
        combined_data.pop('Days')
        combined_data['交割日期'] = combined_data['交割日期'].apply(lambda x: x.strftime('%Y%m%d') if pd.notna(x) else '')
        combined_data['交易日期'] = combined_data['交易日期'].apply(lambda x: x.strftime('%Y%m%d') if pd.notna(x) else '')
        combined_data['成交日期'] = combined_data['成交日期'].apply(lambda x: x.strftime('%Y%m%d') if pd.notna(x) else '')
        # ============================備註內容=================================

        # 設定表格行數
        table_cbas_to_cb.setRowCount(len(combined_data))
        
        # 填入資料
        for i, row in combined_data.iterrows():
            for j, col in enumerate(combined_data.columns):
                item = QTableWidgetItem(str(row[col]))
                
                # 設定表格可編輯
                item.setFlags(item.flags() | 16)  # Qt.ItemIsEditable
                
                # 設定背景色
                if col in ['履約張數', '履約價', '賣出金額']:
                    item.setBackground(QColor(255, 255, 224))  # 淺黃色
                elif col in ['交易日期', '交割日期']:
                    item.setBackground(QColor(204, 229, 255))  # 淺藍色
                
                table_cbas_to_cb.setItem(i, j, item)
        
        QMessageBox.information(None, "成功", f"已新增 {len(upper_data)} 筆履約資訊到本日新增履約表格！")
        
    except Exception as e:
        QMessageBox.critical(None, "新增失敗", f"發生錯誤：{e}")


def add_exercise_to_sell(table_cbas_to_cb, df_quote, get_table_data_func, show_sell_table_func):
    """將實物履約添加到提解賣出分頁"""
    try:
        # 檢查本日新增履約表格是否有資料
        if table_cbas_to_cb.rowCount() == 0:
            QMessageBox.warning(None, "警告", "請先新增履約資訊！")
            return
        
        # 取得本日新增履約表格的資料
        exercise_data = get_table_data_func(table_cbas_to_cb)
        
        if exercise_data.empty:
            QMessageBox.warning(None, "警告", "本日新增履約表格沒有資料！")
            return
        
        # 轉換為提解賣出格式
        sell_columns = [
        '原單契約編號', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '交易日期', '交割日期',
        '履約張數', '成交均價', '履約利率', '提前履約賠償金', '履約價', 
        '選擇權交割單價', '交割總金額', '錄音時間', '解約類別', '履約方式'
            ]

        exercise_data['履約價'] = exercise_data['履約價'].astype(float)
        exercise_data['成交均價'] = exercise_data['履約價']
        exercise_data['履約利率'] = exercise_data['原利率']
        exercise_data['履約張數'] = exercise_data['履約張數'].astype(int)

        exercise_data['選擇權交割單價'] = exercise_data['成交均價'] - exercise_data['履約價']
        exercise_data['交割總金額'] = exercise_data['履約張數'] * exercise_data['履約價'] * 1000
        exercise_data['提前履約賠償金'] = '0'

        df_sell_col = pd.DataFrame(columns=sell_columns)
        df_sell_data = pd.concat([df_sell_col, exercise_data], ignore_index=True)
        df_sell_data['來自'] = '實物履約'
        show_sell_table_func(df_sell_data, from_where='Execution')
        QMessageBox.information(None, "成功", f"已將 {len(df_sell_data)} 筆實物履約添加到提解賣出分頁！")
        
    except Exception as e:
        QMessageBox.critical(None, "新增失敗", f"發生錯誤：{e}")


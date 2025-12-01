import pandas as pd
import pyodbc
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt
from datetime import datetime
import numpy as np
from format_utils import strip_trailing_zeros, next_business_day
from format_utils import strip_whitespace


def query_renewal_contracts(cus_id_text, cb_code_text, df_quote, table_renewal_query):
    """查詢續期合約資料並進行聚合"""
    try:
        # 從ComboBox中提取客戶ID和CB代號
        cus_id = cus_id_text.split(" - ")[0] if " - " in cus_id_text else cus_id_text
        cb_code = cb_code_text.split(" - ")[0] if " - " in cb_code_text else cb_code_text
        
        # 驗證輸入（客戶ID和CB代號至少要輸入一個）
        if not cus_id and not cb_code:
            QMessageBox.warning(None, "輸入錯誤", "請至少輸入客戶ID或CB代號！")
            return None
        
        # 從資料庫查詢合約資料
        conn = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
        
        # 建構查詢條件
        where_conditions = ["a.STORQTY > 0"]
        
        if cus_id:
            # 補齊客戶ID到12位
            db_cusid_length = 12
            cus_id_padded = cus_id.ljust(db_cusid_length)
            where_conditions.append(f"a.CUSID = '{cus_id_padded}'")
        
        if cb_code:
            where_conditions.append(f"a.CBCODE = '{cb_code}'")
        
        where_clause = " AND ".join(where_conditions)
        
        # 保持原SQL不變，取得所有個別契約資料
        sql_query = f"""
            SELECT 
                a.PRDID,
                a.CUSID,
                c.CUSNAME,
                a.CBCODE,
                a.STORQTY,
                a.PERRATE,
                a.CBTCOST,
                a.TRDATE,
                a.CBTPDT,
                a.CBTPPRI,
                a.OPTEXDT
            FROM FSPFLIB.ASPROD a 
            LEFT JOIN FSPFLIB.FSPCS0M c ON a.CUSID = c.CUSID 
            WHERE {where_clause}
            ORDER BY a.PRDID DESC
        """
        
        # 使用pd.read_sql直接取得DataFrame
        df_contracts = pd.read_sql(sql_query, conn)
        conn.close()
        
        if df_contracts.empty:
            QMessageBox.information(None, "查詢結果", "沒有找到符合條件的合約資料！")
            table_renewal_query.setRowCount(0)
            return None
        
        # 保存原始合約資料供後續使用
        df_original_contracts = df_contracts.copy()
        
        df_contracts = strip_whitespace(df_contracts)
        
        # 重新命名欄位為中文
        df_contracts.rename(columns={
            'PRDID': '新作契約編號',
            'CUSID': '客戶ID', 
            'CUSNAME': '客戶名稱',
            'CBCODE': 'CB代號',
            'STORQTY': '原庫存張數',
            'PERRATE': '原履約利率',
            'CBTCOST': '成交均價',
            'TRDATE': '原交易日期',
            'CBTPDT': '賣回日',
            'CBTPPRI': '賣回價',
            'OPTEXDT': '選擇權到期日'
        }, inplace=True)
        
        # 合併CB名稱
        df_contracts['CB代號'] = df_contracts['CB代號'].apply(lambda x: strip_trailing_zeros(x)).astype(str)
        df_contracts = df_contracts.merge(df_quote[['CB代號', 'CB名稱']], on='CB代號', how='left')
        
        # 讀取ASCCSV02.csv取得今日賣出張數
        try:
            tday = datetime.now().strftime('%Y%m%d')
            csv_path_sell = rf"\\10.72.228.83\CBAS_Logs\{tday}\AfterClose\ASCCSV02.csv"
            df_sell = pd.read_csv(csv_path_sell, encoding='utf-8-sig')
            
            # 設定欄位名稱（根據你提供的結構）
            sell_columns_map = {
                df_sell.columns[0]: 'PRDID',   # 契約編號
                df_sell.columns[1]: 'CUSID',   # 客戶ID  
                df_sell.columns[2]: 'CBCODE',  # CB代號
                df_sell.columns[7]: 'DEUQTY'   # 賣出張數 (第8欄，索引7)
            }
            df_sell = df_sell.rename(columns=sell_columns_map)
            
            # 聚合賣出資料
            df_sell_agg = df_sell.groupby(['CUSID', 'CBCODE'])['DEUQTY'].sum().reset_index()
            df_sell_agg.rename(columns={'DEUQTY': '今賣出張數'}, inplace=True)
            # 確保數據類型一致
            df_sell_agg['CUSID'] = df_sell_agg['CUSID'].astype(str).str.strip()
            df_sell_agg['CBCODE'] = df_sell_agg['CBCODE'].apply(lambda x: strip_trailing_zeros(x)).astype(str)
            
        except FileNotFoundError:
            print(f"今日交易檔案不存在：{csv_path_sell}")
            # 如果沒有賣出檔案，今賣出張數設為0
            df_sell_agg = pd.DataFrame(columns=['CUSID', 'CBCODE', '今賣出張數'])
        
        # 聚合合約資料 GROUP BY 客戶ID, 客戶名稱, CB名稱, CB代號
        df_aggregated = df_contracts.groupby(['客戶ID', '客戶名稱', 'CB代號', 'CB名稱']).agg({
            '原庫存張數': 'sum'
        }).reset_index()
        
        # 確保merge欄位的數據類型一致
        df_aggregated['原庫存張數'] = df_aggregated['原庫存張數'].astype(int)
        df_aggregated['客戶ID'] = df_aggregated['客戶ID'].astype(str).str.strip()
        df_aggregated['CB代號'] = df_aggregated['CB代號'].astype(str)
        
        # 合併今日賣出張數
        df_aggregated = df_aggregated.merge(
            df_sell_agg, 
            left_on=['客戶ID', 'CB代號'], 
            right_on=['CUSID', 'CBCODE'], 
            how='left'
        )
        df_aggregated['今賣出張數'] = df_aggregated['今賣出張數'].fillna(0).astype(int)
        df_aggregated.drop(['CUSID', 'CBCODE'], axis=1, errors='ignore', inplace=True)
        
        # 計算今剩餘張數
        df_aggregated['今剩餘張數'] = (df_aggregated['原庫存張數'] - df_aggregated['今賣出張數']).astype(int)
        
        # 添加今履約利率（從報價表取得）
        # 確保兩個DataFrame的CB代號都是字符串類型
        df_aggregated['CB代號'] = df_aggregated['CB代號'].astype(str)
        df_quote_for_merge = df_quote[['CB代號', '履約利率']].copy()
        df_quote_for_merge['CB代號'] = df_quote_for_merge['CB代號'].astype(str)
        
        df_aggregated = df_aggregated.merge(df_quote_for_merge, on='CB代號', how='left')
        df_aggregated.rename(columns={'履約利率': '今履約利率'}, inplace=True)
        df_aggregated['今履約利率'] = pd.to_numeric(df_aggregated['今履約利率'], errors='coerce').fillna(0)
        
        # 新增空白欄位供用戶輸入
        df_aggregated['續期張數'] = ''
        df_aggregated['今成交均價'] = ''
        
        # 確保欄位順序
        final_columns = ['客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '原庫存張數', '今賣出張數', '今剩餘張數', '續期張數', '今履約利率', '今成交均價']
        df_aggregated = df_aggregated[final_columns]
        
        # 更新上方查詢結果表格
        table_renewal_query.setRowCount(len(df_aggregated))
        for i, row in df_aggregated.iterrows():
            # 第一欄：添加checkbox
            checkbox = QTableWidgetItem()
            checkbox.setFlags(checkbox.flags() | Qt.ItemIsUserCheckable)
            checkbox.setCheckState(Qt.Unchecked)
            table_renewal_query.setItem(i, 0, checkbox)
            
            # 其他欄位：從第二欄開始
            for j, col in enumerate(df_aggregated.columns):
                value = str(row[col]) if pd.notna(row[col]) else ""
                item = QTableWidgetItem(value)
                
                # 設定可編輯的欄位
                if col in ['續期張數', '今履約利率', '今成交均價']:
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    item.setBackground(QColor(255, 255, 224))  # 淺黃色表示可編輯

                elif col in ['客戶名稱', 'CB名稱', '今剩餘張數']:
                    item.setBackground(QColor(204, 229, 255))  # 淺綠色
                
                table_renewal_query.setItem(i, j + 1, item)  # j+1 因為第0欄是checkbox
        
        QMessageBox.information(None, "查詢完成", f"找到 {len(df_aggregated)} 筆聚合的客戶+標的組合！")
        
        return df_original_contracts
        
    except Exception as e:
        QMessageBox.critical(None, "查詢失敗", f"發生錯誤：{e}")
        print(f"查詢續期合約錯誤：{e}")
        try:
            if 'conn' in locals() and conn:
                conn.close()
        except:
            pass  # 忽略連接關閉錯誤
        return None


def add_renewal_contract(table_renewal_query, table_renewal_buy, table_renewal_sell, df_original_contracts, df_quote):
    """新增選中的續期合約到左右兩側表格"""
    try:
        # 檢查上方查詢結果表格是否有資料
        if table_renewal_query.rowCount() == 0:
            QMessageBox.warning(None, "警告", "請先查詢續期合約資料！")
            return
        
        if df_original_contracts is None or df_original_contracts.empty:
            QMessageBox.warning(None, "警告", "請重新執行查詢！")
            return
        
        # 收集被勾選的聚合資料
        selected_aggregated = []
        table = table_renewal_query
        
        for row in range(table.rowCount()):
            # 檢查第0欄的checkbox是否被勾選
            checkbox_item = table.item(row, 0)
            if checkbox_item and checkbox_item.checkState() == Qt.Checked:
                # 收集該行資料（跳過第0欄的checkbox）
                row_data = {}
                for col in range(1, table.columnCount()):  # 從第1欄開始，跳過checkbox
                    item = table.item(row, col)
                    col_name = table.horizontalHeaderItem(col).text()
                    row_data[col_name] = item.text() if item else ""
                selected_aggregated.append(row_data)
        
        if not selected_aggregated:
            QMessageBox.warning(None, "警告", "請先勾選要新增的合約！")
            return
        
        # 1. 處理左側新作表格 - 直接加入聚合資料
        new_columns = ['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '續期張數', '今履約利率', '今成交均價']
        new_data = []
        for agg_row in selected_aggregated:
            new_row = {col: agg_row.get(col, '') for col in new_columns}
            new_data.append(new_row)
        
        # 取得左側表格現有資料
        current_new_data = get_table_data(table_renewal_buy)
        if not current_new_data.empty:
            new_df = pd.concat([current_new_data, pd.DataFrame(new_data)], ignore_index=True)
        else:
            new_df = pd.DataFrame(new_data)
        
        # 更新左側新作表格
        update_renewal_table(table_renewal_buy, new_df, new_columns)
        
        # 2. 處理右側賣出表格 - 按續期張數累計契約明細
        sell_data = []
        for agg_row in selected_aggregated:
            cus_id = agg_row['客戶ID']
            cb_code = agg_row['CB代號']
            renewal_qty_str = agg_row.get('續期張數', '0')
            today_price = agg_row.get('今成交均價', '0')
            
            # 將續期張數轉為整數
            try:
                renewal_qty = int(renewal_qty_str) if renewal_qty_str else 0
            except:
                renewal_qty = 0
            
            if renewal_qty <= 0:
                continue  # 跳過沒有續期張數的項目
            
            # 從原始合約資料中找到該客戶該標的的所有契約
            cus_contracts = df_original_contracts[
                (df_original_contracts['CUSID'].astype(str).str.strip() == str(cus_id).strip()) & 
                (df_original_contracts['CBCODE'].apply(lambda x: strip_trailing_zeros(x)).astype(str) == str(cb_code))
            ].copy()
            
            if cus_contracts.empty:
                continue
            
            # 按原交易日期排序（越早越先續期）
            cus_contracts = cus_contracts.sort_values('TRDATE')
            
            # 讀取ASCCSV02.csv檢查已賣出張數
            try:
                tday = datetime.now().strftime('%Y%m%d')
                csv_path_sell = rf"\\10.72.228.83\CBAS_Logs\{tday}\AfterClose\ASCCSV02.csv"
                df_sell = pd.read_csv(csv_path_sell, encoding='utf-8-sig')
                # 使用用戶修正的欄位位置
                prdid_col = df_sell.columns[0]
                qty_col = df_sell.columns[7]  # 用戶修正為第8欄（索引7）
                df_sell_lookup = df_sell.set_index(prdid_col)[qty_col].to_dict()
            except:
                df_sell_lookup = {}
            
            # 從報價表取得CB名稱
            cb_info = df_quote[df_quote['CB代號'].astype(str) == str(cb_code)]
            cb_name = cb_info.iloc[0]['CB名稱'] if not cb_info.empty else ''
            
            # 按順序累計契約，直到達到續期張數
            accumulated_qty = 0
            for _, contract in cus_contracts.iterrows():
                if accumulated_qty >= renewal_qty:
                    break  # 已達到需要的續期張數
                
                prdid = contract['PRDID']
                storqty = contract['STORQTY']
                sold_qty = df_sell_lookup.get(prdid, 0)
                remaining_qty = storqty - sold_qty
                
                # 只處理有剩餘張數的契約
                if remaining_qty > 0:
                    # 計算這個契約要續期的張數
                    needed_qty = renewal_qty - accumulated_qty
                    contract_renewal_qty = min(remaining_qty, needed_qty)
                    
                    # 建立賣出明細行
                    sell_row = {
                        '新作契約編號': prdid,
                        '客戶ID': cus_id,
                        '客戶名稱': agg_row['客戶名稱'],
                        'CB代號': cb_code,
                        'CB名稱': cb_name,
                        '原庫存張數': storqty,
                        '今賣出張數': sold_qty,
                        '續期張數': contract_renewal_qty,
                        '成交均價': today_price
                    }
                    sell_data.append(sell_row)
                    
                    # 累計已處理的張數
                    accumulated_qty += contract_renewal_qty
        
        # 取得右側表格現有資料
        current_sell_data = get_table_data(table_renewal_sell)
        if not current_sell_data.empty:
            sell_df = pd.concat([current_sell_data, pd.DataFrame(sell_data)], ignore_index=True)
        else:
            sell_df = pd.DataFrame(sell_data)
        
        # 更新右側賣出表格
        sell_columns = ['新作契約編號', '客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '原庫存張數', '今賣出張數', '續期張數', '成交均價']
        update_renewal_table(table_renewal_sell, sell_df, sell_columns)
        
        QMessageBox.information(None, "新增成功", f"已新增 {len(selected_aggregated)} 筆聚合資料到新作表格，{len(sell_data)} 筆契約明細到賣出表格！")
        
    except Exception as e:
        QMessageBox.critical(None, "新增失敗", f"發生錯誤：{e}")
        print(f"新增續期合約錯誤：{e}")
        import traceback
        traceback.print_exc()


def update_renewal_table(table, df, columns):
    """通用方法：用DataFrame更新表格"""
    if df.empty:
        table.setRowCount(0)
        return
        
    table.setRowCount(len(df))
    for i, row in df.iterrows():
        for j, col in enumerate(columns):
            value = str(row.get(col, '')) if pd.notna(row.get(col, '')) else ""
            item = QTableWidgetItem(value)
            
            # 設定背景色
            if col in ['續期張數', '今履約利率', '今成交均價', '成交均價']:
                item.setBackground(QColor(255, 255, 224))  # 淺黃色
            elif col in ['客戶名稱', 'CB名稱']:
                item.setBackground(QColor(204, 229, 255))  # 淺綠色
            elif col in ['新作契約編號', '原庫存張數']:
                item.setBackground(QColor(204, 229, 255))  # 淺藍色
            
            table.setItem(i, j, item)


def transfer_renewal_data(table_renewal_buy, table_renewal_sell, df_original_contracts, 
                         calculate_new_trade_batch, show_buy_table, show_sell_table):
    """將續期合約資料轉換為交易資料"""
    try:
        # 1. 處理左側買進資料
        buy_data = get_table_data(table_renewal_buy)
        if not buy_data.empty:
            # 添加calculate_new_trade_batch需要的欄位
            tday = datetime.now()
            settle_date = next_business_day(tday, 2)
            
            # 重新命名和添加必要欄位
            buy_data = buy_data.rename(columns={
                '續期張數': '成交張數',
                '今成交均價': '成交均價'
            })
            
            # 添加必要的日期欄位
            buy_data['交易日期'] = tday.strftime('%Y%m%d')
            buy_data['交割日期'] = settle_date.strftime('%Y%m%d')
            buy_data['錄音時間'] = ''
            buy_data['交易類型'] = 'ASO'
            buy_data['成交金額'] = round(buy_data['成交張數'].astype(int) * buy_data['成交均價'].astype(float) * 1000, 0).astype(int)
            buy_data['來自'] = '續期'
            # 使用calculate_new_trade_batch處理買進資料
            processed_buy_data = calculate_new_trade_batch(buy_data)
            
            # 使用show_buy_table顯示買進資料
            show_buy_table(processed_buy_data)
            QMessageBox.information(None, "成功", f"已處理 {len(processed_buy_data)} 筆買進資料！")
        
        # 2. 處理右側賣出資料
        sell_data = get_table_data(table_renewal_sell)
        if not sell_data.empty:
            # 與原始合約資料合併以取得詳細資訊
            if df_original_contracts is not None and not df_original_contracts.empty:
                sell_data = sell_data.merge(
                    df_original_contracts, 
                    left_on='新作契約編號', 
                    right_on='PRDID', 
                    how='left'
                )
            
            # 重新命名欄位
            sell_data.rename(columns={
                '新作契約編號': '原單契約編號',
                'CBTPDT': '賣回日',
                'CBTPPRI': '賣回價',
                'OPTEXDT': '選擇權到期日',
                'PERRATE': '履約利率'
            }, inplace=True)
            
            # 新增必要欄位
            tday = datetime.now()
            settlement_date = next_business_day(tday, 2)
            sell_data['交易日期'] = tday.strftime('%Y%m%d')
            sell_data['交割日期'] = settlement_date.strftime('%Y%m%d')
            sell_data['履約張數'] = pd.to_numeric(sell_data['續期張數'], errors='coerce').fillna(0).astype(int)
            sell_data['剩餘張數'] = pd.to_numeric(sell_data['原庫存張數'], errors='coerce') - pd.to_numeric(sell_data['今賣出張數'], errors='coerce') - sell_data['履約張數'] #用來判斷解約類別為何
            sell_data['解約類別'] = np.where(sell_data['剩餘張數'] > 0, '1', '2')  # 2 = 全, 1 = 部分
            sell_data['履約方式'] = '1'   # 1 = 現金履約
            sell_data['提前履約賠償金'] = '0'
            sell_data['錄音時間'] = ''
            
            # 計算履約價（需要從賣回價和利率計算）
            sell_data['年期'] = ((pd.to_datetime(sell_data['賣回日'], format='%Y%m%d') - pd.to_datetime(sell_data['交割日期'], format='%Y%m%d')).dt.days + 1) / 365
            sell_data['履約價'] = round(pd.to_numeric(sell_data['賣回價'], errors='coerce').fillna(0) - sell_data['年期'] * pd.to_numeric(sell_data['履約利率'], errors='coerce').fillna(0), 2)
            sell_data['選擇權交割單價'] = pd.to_numeric(sell_data['成交均價'], errors='coerce') - sell_data['履約價']
            sell_data['交割總金額'] = sell_data['履約張數'] * sell_data['選擇權交割單價']
            sell_data['來自'] = '續期'
            # 使用show_sell_table顯示賣出資料
            show_sell_table(sell_data, from_where='Renewal')
            QMessageBox.information(None, "成功", f"已處理 {len(sell_data)} 筆賣出資料！")
        
        if buy_data.empty and sell_data.empty:
            QMessageBox.warning(None, "警告", "請先新增續期合約資料！")
            
    except Exception as e:
        QMessageBox.critical(None, "轉換失敗", f"發生錯誤：{e}")
        print(f"轉換續期資料錯誤：{e}")
        import traceback
        traceback.print_exc()


def get_table_data(table):
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
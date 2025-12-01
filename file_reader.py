import pandas as pd
import os
from format_utils import strip_trailing_zeros
from db_access import get_cbas_customers


def read_quote_excel(file_path: str = None) -> pd.DataFrame:
    """讀取CBAS報價表Excel檔案"""
    if file_path is None:
        file_path = r"\\10.72.228.112\cbas業務公用區\統一證CBAS報價表_內部.xlsm"
    
    try:
        df_quote = pd.read_excel(file_path, sheet_name='aso報價', header=[2], usecols='A:AI')
        df_quote = df_quote[['CB代號', 'CB名稱', '選擇權到期日', '賣回日', '賣回價', '百元報價', '履約利率', '低百元報價', '低履約利率', '波動度']]

        # 過濾掉"元富專用報價"以下的資料
        if not df_quote[df_quote['CB名稱'] == '元富專用報價'].empty:
            cutoff_idx = df_quote[df_quote['CB名稱'] == '元富專用報價'].index[0]
            df_quote = df_quote.iloc[:cutoff_idx].reset_index(drop=True)
        df_quote = df_quote.dropna().reset_index(drop=True)

        # 處理日期格式：轉換為YYYYMMDD格式
        def format_date_to_yyyymmdd(date_val):
            try:
                if pd.isna(date_val):
                    return ''
                if isinstance(date_val, str):
                    if len(date_val) == 8 and date_val.isdigit():
                        return date_val  # 已經是YYYYMMDD格式
                    date_val = pd.to_datetime(date_val)
                elif isinstance(date_val, (int, float)):
                    date_val = pd.to_datetime(date_val, origin='1899-12-30', unit='D')
                else:
                    date_val = pd.to_datetime(date_val)
                return date_val.strftime('%Y%m%d')
            except:
                return str(date_val) if not pd.isna(date_val) else ''

        df_quote['選擇權到期日'] = df_quote['選擇權到期日'].apply(format_date_to_yyyymmdd)
        df_quote['賣回日'] = df_quote['賣回日'].apply(format_date_to_yyyymmdd)

        # 將履約利率乘以100並保留兩位小數
        df_quote['履約利率'] = df_quote['履約利率'].apply(lambda x: round(float(x)*100, 2))
        df_quote['低履約利率'] = df_quote['低履約利率'].apply(lambda x: round(float(x)*100, 2))

        numeric_columns = ['百元報價', '低百元報價', '波動度']
        for col in numeric_columns:
            df_quote[col] = df_quote[col].round(2)

        df_quote['CB代號'] = df_quote['CB代號'].apply(strip_trailing_zeros).astype(str)

        # 去除完全重複的記錄，保留第一筆
        df_quote = df_quote.drop_duplicates(keep='first')
        
        # 檢查重複的CB代號
        duplicate_cb = df_quote[df_quote.duplicated(subset=['CB代號'], keep=False)]
        if not duplicate_cb.empty:
            duplicate_cb_list = duplicate_cb['CB代號'].unique().tolist()
            duplicate_msg = f"警告：報價表發現重複的CB代號：{', '.join(duplicate_cb_list)}"
            print(duplicate_msg)

        return df_quote
    except Exception as e:
        print(f"讀取報價表時發生錯誤: {e}")
        return pd.DataFrame(columns=['CB代號', 'CB名稱', '選擇權到期日', '賣回日', '賣回價', '百元報價', '履約利率', '低百元報價', '低履約利率', '波動度'])

def read_vip_list(file_path: str = None) -> pd.DataFrame:
    """讀取VIP名單CSV檔案"""
    if file_path is None:
        file_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\VIP_List.csv"
    
    try:
        # 嘗試多種編碼讀取
        for encoding in ['utf-8-sig', 'utf-8', 'big5', 'gbk', 'cp950']:
            try:
                df_vip_list = pd.read_csv(file_path, encoding=encoding)
                return df_vip_list
            except (UnicodeDecodeError, FileNotFoundError):
                continue
        
        print(f"無法讀取 {file_path}，將創建空的 VIP 名單")
        return pd.DataFrame(columns=['客戶ID', '客戶名稱', '不限張數低手續費', '不限張數低利率'])
    except Exception as e:
        print(f"讀取VIP名單時發生錯誤: {e}")
        return pd.DataFrame(columns=['客戶ID', '客戶名稱', '不限張數低手續費', '不限張數低利率'])

def read_vip_quote(file_path: str = None) -> pd.DataFrame:
    """讀取VIP特殊報價CSV檔案"""
    if file_path is None:
        file_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\VIP_Quote.csv"
    
    try:
        # 嘗試多種編碼讀取
        for encoding in ['utf-8-sig', 'utf-8', 'big5', 'gbk', 'cp950']:
            try:
                df_vip_quote = pd.read_csv(file_path, encoding=encoding)
                return df_vip_quote
            except (UnicodeDecodeError, FileNotFoundError):
                continue
        
        print(f"無法讀取 {file_path}，將創建空的 VIP 報價")
        return pd.DataFrame(columns=['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '利率%', '手續費'])
    except Exception as e:
        print(f"讀取VIP報價時發生錯誤: {e}")
        return pd.DataFrame(columns=['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '利率%', '手續費'])

def read_today_trade_buy(tday, settle_date, csv_path: str = None) -> pd.DataFrame:
        tdaystr = tday.strftime('%Y%m%d')
        csv_path_buy = rf"\\10.72.228.83\CBAS_Logs\{tdaystr}\AfterClose\BuyMatch.csv"
        #csv_path_buy = r"\\10.72.228.83\CBAS_Logs\20250909\AfterClose\BuyMatch.csv"
        df_csv_buy = pd.read_csv(csv_path_buy, encoding='utf-8-sig')  # 使用 utf-8-sig 以支援 BOM
        
        df_csv_buy_groupby = df_csv_buy.groupby(['CUSID', 'CBCODE', 'SRC'], dropna=False).agg({
            'MATCHQTY': 'sum',
            'MATCHAMT': 'sum'
        }).reset_index()

        # 客戶名稱查詢時，確保CUSID格式一致
        df_cusname = get_cbas_customers()
        
        df_buy = df_csv_buy_groupby.merge(df_cusname, left_on='CUSID', right_on='CUSID', how='left', suffixes=('', '_name'))
        df_buy['CBCODE'] = df_buy['CBCODE'].apply(lambda x: strip_trailing_zeros(x)).astype(str)
        df_buy['成交張數'] = (df_buy['MATCHQTY'] / 1000).astype(int)
        df_buy['成交均價'] = df_buy['MATCHAMT'] / df_buy['MATCHQTY']
        df_buy = df_buy.rename(columns={
            'CUSID': '客戶ID',
            'CBCODE': 'CB代號',
            'CUSNAME': '客戶名稱',
            'MATCHAMT': '成交金額'
        })


        df_buy['交易類型'] = 'ASO'
        df_buy['交割日期'] = settle_date.strftime('%Y%m%d')
        df_buy['來自'] = '盤面交易'
        return df_buy

def read_today_trade_sell(tday, settle_date, csv_path: str = None) -> pd.DataFrame:
        tdaystr = tday.strftime('%Y%m%d')
        csv_path_sell = rf"\\10.72.228.83\CBAS_Logs\{tdaystr}\AfterClose\ASCCSV02.csv"
        #csv_path_sell = r"\\10.72.228.83\CBAS_Logs\20250909\AfterClose\ASCCSV02.csv"
        df_csv_sell = pd.read_csv(csv_path_sell, encoding='utf-8-sig')  # 使用 utf-8-sig 以支援 BOM
        
        # 你指定的欄位名稱（不含解約契約編號）
        user_columns_sell = [
            '原單契約編號', '客戶ID', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
            '履約張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日', '提前履約賠償金', '履約價', '選擇權交割單價', '交割總金額', '錄音時間'
        ]
        df_csv_sell.columns = user_columns_sell  # 只改欄位名稱，不動內容
        df_csv_sell['來自'] = '盤面交易'

        return df_csv_sell

def read_customer_list(file_path: str = None) -> pd.DataFrame:
    """讀取客戶清單CSV檔案"""
    if file_path is None:
        file_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\Customer_List.csv"
    
    try:
        for encoding in ['utf-8-sig', 'utf-8', 'big5', 'gbk', 'cp950']:
            try:
                df_customer = pd.read_csv(file_path, encoding=encoding)
                return df_customer
            except (UnicodeDecodeError, FileNotFoundError):
                continue
        
        print(f"無法讀取 {file_path}，將創建空的客戶清單")
        return pd.DataFrame(columns=['客戶ID', '客戶名稱'])
    except Exception as e:
        print(f"讀取客戶清單時發生錯誤: {e}")
        return pd.DataFrame(columns=['客戶ID', '客戶名稱'])

def load_vip_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    """讀取VIP資料並返回兩個DataFrame"""
    vip_list_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\VIP_List.csv"
    vip_quote_path = r"\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\VIP_Quote.csv"
    
    # 嘗試多種編碼讀取 VIP_List.csv
    df_vip_list = None
    for encoding in ['utf-8-sig', 'utf-8', 'big5', 'gbk', 'cp950']:
        try:
            df_vip_list = pd.read_csv(vip_list_path, encoding=encoding)
            break
        except (UnicodeDecodeError, FileNotFoundError):
            continue
    
    # 嘗試多種編碼讀取 VIP_Quote.csv
    df_vip_quote = None
    for encoding in ['utf-8-sig', 'utf-8', 'big5', 'gbk', 'cp950']:
        try:
            df_vip_quote = pd.read_csv(vip_quote_path, encoding=encoding)
            break
        except (UnicodeDecodeError, FileNotFoundError):
            continue
    
    # 如果檔案不存在，創建空的 DataFrame
    if df_vip_list is None:
        print(f"無法讀取 {vip_list_path}，將創建空的 VIP 名單")
        df_vip_list = pd.DataFrame(columns=['客戶ID', '客戶名稱', '不限張數低手續費', '不限張數低利率'])
        
    if df_vip_quote is None:
        print(f"無法讀取 {vip_quote_path}，將創建空的 VIP 報價")
        df_vip_quote = pd.DataFrame(columns=['客戶ID', '客戶名稱', 'CB代號', 'CB名稱', '利率%', '手續費'])
    
    return df_vip_list, df_vip_quote

def get_daily_bond_rate() -> str:
    """讀取日指標公債利率，返回殖利率字串"""
    from WCFAdox import PCAX
    from datetime import datetime, timedelta
    from format_utils import strip_trailing_zeros
    
    PX = PCAX("10.72.241.51")
    yes = pd.Timestamp.today() - pd.Timedelta(days=1)
    yesstr = yes.strftime('%Y%m%d')
    
    try:
        df_rf = PX.Sil_Data("日指標公債利率", "D", "RA05", yesstr, yesstr, isst="N")
        while len(df_rf) == 0:
            yes = yes - pd.Timedelta(days=1)
            yesstr = yes.strftime('%Y%m%d')
            df_rf = PX.Sil_Data("日指標公債利率", "D", "RA05", yesstr, yesstr, isst="N")
        
        rf = strip_trailing_zeros(df_rf['殖利率(%)'][0])
        return rf
    except Exception as e:
        print(f"讀取日指標公債利率發生錯誤: {e}")
        return '1.45' 

def read_expired_trade_data():
    """讀取指定日期的到期交易資料"""
    from datetime import datetime
    try:
        tdaystr = datetime.now().strftime('%Y%m%d')
        csv_path_sell = rf"\\10.72.228.83\CBAS_Logs\{tdaystr}\AfterClose\ASCCSV02.csv"
        df_sell = pd.read_csv(csv_path_sell, encoding='utf-8-sig')

        user_columns_sell = [
            '原單契約編號', '客戶ID', 'CB代號', '交易日期', '交割日期', '解約類別', '履約方式',
            '履約張數', '成交均價', '履約利率', '賣回日', '賣回價', '選擇權到期日', '提前履約賠償金', '履約價', '選擇權交割單價', '交割總金額', '錄音時間'
        ]
        df_sell.columns = user_columns_sell  # 只改欄位名稱，不動內容
        #df_sell['選擇權到期日'] = df_sell['選擇權到期日'].astype(str)

        # 篩選指定日期到期的交易，並按原單契約編號統計已賣出張數
        #df_sell_today = df_sell[df_sell['選擇權到期日'] == target_date]
        df_sell_summary = df_sell.groupby('原單契約編號').agg({'履約張數': 'sum'}).reset_index()
        df_sell_summary.rename(columns={'履約張數': '今日賣出張數_ASCCSV02'}, inplace=True)

        return df_sell_summary

    except FileNotFoundError:
        print(f"指定日期交易檔案不存在：{csv_path_sell}")
        # 如果沒有交易檔案，創建空的DataFrame
        df_sell_summary = pd.DataFrame(columns=['原單契約編號', '今日賣出張數_ASCCSV02'])

        return df_sell_summary
    except Exception as e:
        print(f"讀取到期交易資料時發生錯誤: {e}")
        df_sell_summary = pd.DataFrame(columns=['原單契約編號', '今日賣出張數_ASCCSV02'])

        return df_sell_summary
    
def load_quote(): #讀取報價表
        #df_cbinfo = pd.read_excel(r'\\10.72.228.112\cbas業務公用區\統一證CBAS報價表_內部.xlsm', sheet_name='彙整CB基本資料', header=[4])
        df_quote = pd.read_excel(r'\\10.72.228.112\cbas業務公用區\統一證CBAS報價表_內部.xlsm', sheet_name='aso報價', header=[2], usecols='A:AI')
        df_quote = df_quote[['CB代號', 'CB名稱', '選擇權到期日', '賣回日', '賣回價', '百元報價', '履約利率', '低百元報價', '低履約利率', '波動度']]

        # 過濾掉"元富專用報價"以下的資料
        if not df_quote[df_quote['CB名稱'] == '元富專用報價'].empty:
            cutoff_idx = df_quote[df_quote['CB名稱'] == '元富專用報價'].index[0]
            df_quote = df_quote.iloc[:cutoff_idx].reset_index(drop=True)
        df_quote = df_quote.dropna().reset_index(drop=True)

        # 處理日期格式：轉換為YYYYMMDD格式
        def format_date_to_yyyymmdd(date_val):
            try:
                if pd.isna(date_val):
                    return ''
                if isinstance(date_val, str):
                    # 如果已經是字符串，嘗試解析
                    if len(date_val) == 8 and date_val.isdigit():
                        return date_val  # 已經是YYYYMMDD格式
                    date_val = pd.to_datetime(date_val)
                elif isinstance(date_val, (int, float)):
                    # 如果是數字，可能是Excel的序列日期
                    date_val = pd.to_datetime(date_val, origin='1899-12-30', unit='D')
                else:
                    # 其他情況，直接轉換
                    date_val = pd.to_datetime(date_val)
                
                # 轉換為YYYYMMDD格式
                return date_val.strftime('%Y%m%d')
            except:
                return str(date_val) if not pd.isna(date_val) else ''

        df_quote['選擇權到期日'] = df_quote['選擇權到期日'].apply(format_date_to_yyyymmdd)
        df_quote['賣回日'] = df_quote['賣回日'].apply(format_date_to_yyyymmdd)

        # 將履約利率乘以100並保留兩位小數
        df_quote['履約利率'] = df_quote['履約利率'].apply(lambda x: round(float(x)*100, 2))
        df_quote['低履約利率'] = df_quote['低履約利率'].apply(lambda x: round(float(x)*100, 2))

        numeric_columns = ['百元報價', '低百元報價', '波動度']  # 移除已處理的利率欄位
        for col in numeric_columns:
            df_quote[col] = df_quote[col].round(2)

        df_quote['CB代號'] = df_quote['CB代號'].apply(strip_trailing_zeros).astype(str)

        # 去除完全重複的記錄，保留第一筆
        df_quote = df_quote.drop_duplicates(keep='first')
        # 檢查重複的CB代號
        duplicate_cb = df_quote[df_quote.duplicated(subset=['CB代號'], keep=False)]
        df_cbinfo = pd.read_excel(r'\\10.72.228.112\cbas業務公用區\統一證CBAS報價表_內部.xlsm', sheet_name='彙整CB基本資料', header=[4])
        df_cbinfo = df_cbinfo[['CB代號', 'CB名稱']]
        df_cbinfo['CB代號'] = df_cbinfo['CB代號'].apply(strip_trailing_zeros).astype(str)
        
        return df_quote, duplicate_cb, df_cbinfo

def save_trading_statement(df_bargaining):
    df_bargaining = df_bargaining[['成交日期', '交割日期', 'T+?交割', '錄音時間', '單據編號', '買/賣', '客戶ID', '客戶名稱', 'CB名稱', 'CB代號', '議價張數', '議價價格', '議價金額', '參考價', '備註']]
    df_bargaining['T+?交割'] = df_bargaining['T+?交割'].astype(str)
    df_bargaining['備註二'] = df_bargaining['客戶名稱'] + 'T+' + df_bargaining['T+?交割']
    df = pd.read_excel(r'\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\議價明細.xlsx')
    df = pd.concat([df, df_bargaining])
    df.to_excel(r'\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\議價明細.xlsx', index=False)

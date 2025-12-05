import pandas as pd
import pyodbc
from format_utils import strip_whitespace, next_business_day, prev_business_day
from datetime import datetime
import numpy as np

def get_400_conn():
    conn = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
    return conn

def get_631_conn():
    conn = pyodbc.connect(
        driver='ODBC Driver 18 for SQL Server',
        server='10.72.228.139',
        user='sa',
        password='Self@pscnet',
        database='CBAS',
        TrustServerCertificate='yes'
    )
    return conn

def get_customer_info(cusid_list_padded: list[str]) -> pd.DataFrame:
    """Fetch customer info for a list of padded CUSID (12-char).
    Returns columns: CUSID, CUSNAME, BNKNAME, BNKBRH, BNKACTNO, CENTERNO, ADDRESS2
    """
    if not cusid_list_padded:
        return pd.DataFrame(columns=[
            'CUSID', 'CUSNAME', 'BNKNAME', 'BNKBRH', 'BNKACTNO', 'CENTERNO', 'ADDRESS2'
        ])
    conn = None
    try:
        conn = get_400_conn()
        cusid_list = "','".join(cusid_list_padded)
        sql_query = (
            "SELECT CUSID, CUSNAME, BNKNAME, BNKBRH, BNKACTNO, CENTERNO, ADDRESS2 "
            f"FROM FSPFLIB.FSPCS0M WHERE CBASCODE = 'Y' AND CUSID IN ('{cusid_list}')"
        )
        df = pd.read_sql(sql_query, conn)
        return strip_whitespace(df)
    except Exception:
        return pd.DataFrame(columns=[
            'CUSID', 'CUSNAME', 'BNKNAME', 'BNKBRH', 'BNKACTNO', 'CENTERNO', 'ADDRESS2'
        ])
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass

def get_customer_inventory() -> pd.DataFrame:
    """Fetch customer inventory data from database.
    Returns columns: CUSID, STORQTY
    """
    conn = None
    try:
        conn = get_400_conn()
        df_cus_inventory = strip_whitespace(pd.read_sql(
            "SELECT CUSID, SUM(STORQTY) as STORQTY FROM FSPFLIB.ASPROD GROUP BY CUSID", 
            conn
        ))
        
        if df_cus_inventory.empty:
            print("警告：客戶庫存查詢結果為空")
            return pd.DataFrame(columns=['CUSID', 'STORQTY'])
        
        if 'STORQTY' not in df_cus_inventory.columns:
            print("警告：客戶庫存資料中沒有 STORQTY 欄位")
            df_cus_inventory['STORQTY'] = 0
        
        return df_cus_inventory
        
    except Exception as e:
        print(f"查詢客戶庫存時發生錯誤: {e}")
        return pd.DataFrame()

def get_expired_contracts_db(target_date: str) -> pd.DataFrame:
    """從資料庫取得指定日期到期的契約及客戶名稱"""
    try:
        conn = get_400_conn()

        # 取得指定日期到期的契約及客戶名稱（一次JOIN查詢）
        df_expired = strip_whitespace(pd.read_sql(f"""
            SELECT a.*, c.CUSNAME
            FROM FSPFLIB.ASPROD a
            LEFT JOIN FSPFLIB.FSPCS0M c ON a.CUSID = c.CUSID
            WHERE a.STORQTY > 0 AND a.OPTEXDT = '{target_date}'
        """, conn))

        conn.close()
        return df_expired
    except Exception as e:
        print(f"取得到期契約時發生錯誤: {e}")
        return pd.DataFrame()

def get_cbas_customers() -> pd.DataFrame:
    """Fetch CBAS customers from database.
    Returns columns: CUSID, CUSNAME
    """
    conn = None
    try:
        conn = get_400_conn()
        df_cusname = strip_whitespace(pd.read_sql(
            "SELECT CUSID, CUSNAME FROM FSPFLIB.FSPCS0M WHERE CBASCODE = 'Y'", 
            conn
        ))
        return df_cusname
    except Exception as e:
        print(f"讀取CBAS客戶時發生錯誤: {e}")
        return pd.DataFrame(columns=['CUSID', 'CUSNAME'])
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass

def get_contracts_from_sell_table(df_sell: pd.DataFrame) -> pd.DataFrame:
    """從賣出表格中取得合約資料"""
    try:
        conn = get_400_conn()
        contracts = df_sell['原單契約編號'].unique().tolist()
        contracts_str = "','".join(contracts)
        df_contracts = strip_whitespace(pd.read_sql(
            f"SELECT PRDID, CBTUPRM, OPTTYPE, QPRICE FROM FSPFLIB.ASPROD WHERE PRDID IN ('{contracts_str}')", 
            conn
        ))
        df_contracts.rename(columns={'PRDID': '原單契約編號', 'CBTUPRM': '原單位權利金', 'OPTTYPE': '選擇權型態', 'QPRICE': '報價方式'}, inplace=True)
        return df_contracts
    except Exception as e:
        print(f"從賣出表格中取得合約資料時發生錯誤: {e}")
        return pd.DataFrame()

def get_631_Monitor_Fill():
    conn = get_631_conn()
    df_monitor_fill = pd.read_sql(
        "SELECT * FROM dbo.RPT_Monitor_Fill", 
        conn
    )
    return df_monitor_fill

def get_customer_bank_and_email(cusid_list: list[str]) -> pd.DataFrame:
    """Fetch customer info for a list of padded CUSID (12-char).
    Returns columns: CUSID, CUSNAME, BNKNAME, BNKBRH, BNKACTNO, CENTERNO, ADDRESS2, CELLPHONE
    """
    cusid_list_padded = [cusid.ljust(12) for cusid in cusid_list]
    conn = get_400_conn()
    cusid_list = "','".join(cusid_list_padded)
    sql_query = (
        "SELECT CUSID, CUSNAME, BNKNAME, BNKBRH, BNKACTNO, CENTERNO, EMAIL, CELLPHONE "
        f"FROM FSPFLIB.FSPCS0M WHERE CBASCODE = 'Y' AND CUSID IN ('{cusid_list}')"
    )
    df = pd.read_sql(sql_query, conn)
    return strip_whitespace(df)
  
def get_trust_info(cusid_list: list[str]):
    cusid_list_padded = [cusid.ljust(12) for cusid in cusid_list]
    cusid_list = "','".join(cusid_list_padded)
    conn = get_400_conn()
    df_trust_info = pd.read_sql(
        f"SELECT CUSID, TRUSTEE, TRUSTNM, TRUSTTEL FROM FSPFLIB.FSPCS1M WHERE TRUTYPE = 'T' AND CUSID in ('{cusid_list}')", 
        conn
    )
    return df_trust_info

def get_clearing_detail(tday):
    """
    取得清算明細資料
    
    Args:
        tday: 目標日期 (datetime 物件)
    
    Returns:
        tuple: (df_buy_sum, df_buy_bargain_sum, df_sell_sum, tday_plus_1, tday_plus_2, tday_minus_1, tday_minus_2)
        - df_buy_sum: 買進契約彙總 (客戶ID, 交割日, 權利金總額)
        - df_buy_bargain_sum: 議價交易彙總 (客戶ID, 交割日, 調整後金額)
        - df_sell_sum: 賣出契約彙總 (客戶ID, 到期付款日, 交割總金額)
        - tday_plus_1/2: 目標日期後1/2個工作日
        - tday_minus_1/2: 目標日期前1/2個工作日
    """
    conn = None
    try:
        # 計算相關日期
        tday_plus_1 = next_business_day(tday, 1).strftime("%Y%m%d")
        tday_plus_2 = next_business_day(tday, 2).strftime("%Y%m%d")

        tday_str = tday.strftime("%Y%m%d")
        
        conn = get_400_conn()
        
        # 1. 取得買進契約資料 (ASO類型，過去3個工作日的交易)
        clearing_dates = f"'{tday_plus_1}', '{tday_plus_2}'"
        df_buy_contracts = pd.read_sql(f"""
            SELECT CUSID, SETDATE, PREMTOT 
            FROM FSPFLIB.ASPROD 
            WHERE TXTYPE = 'ASO' 
            AND SETDATE IN ({clearing_dates})
        """, conn)
        
        # 2. 取得賣出契約資料 (未來3個工作日的到期)
        clearing_dates = f"'{tday_plus_1}', '{tday_plus_2}'"
        df_sell_contracts = pd.read_sql(f"""
            SELECT CUSID, PRDID, DUEPAYDT, SETTTOT, CANMODE FROM FSPFLIB.ASSURR 
            WHERE DUEPAYDT IN ({clearing_dates}) OR (DUEDATE = '{tday_str}' AND DUEPAYDT = '{tday_str}')
        """, conn)
        
        
        # 4. 取得議價交易資料
        bargain_dates = f"'{tday_str}', '{tday_plus_1}', '{tday_plus_2}'"
        df_buy_bargain = pd.read_sql(f"""
            SELECT CUSID, SETDAT, TXBS, MTHAMT 
            FROM FSPFLIB.ASBARG 
            WHERE SETDAT IN ({bargain_dates}) AND TXDATE = '{tday_str}'
        """, conn)
        
        # 5. 彙總買進契約資料
        if not df_buy_contracts.empty:
            df_buy_sum = df_buy_contracts.groupby(['CUSID', 'SETDATE']).agg({'PREMTOT': 'sum'}).reset_index()
            df_buy_sum = strip_whitespace(df_buy_sum)
        else:
            df_buy_sum = pd.DataFrame(columns=['CUSID', 'SETDATE', 'PREMTOT'])
        
        # 6. 彙總賣出契約資料
        if not df_sell_contracts.empty:
            df_sell_contracts['SETTTOT'] = np.where(df_sell_contracts['CANMODE'] == '2', -1 * df_sell_contracts['SETTTOT'], df_sell_contracts['SETTTOT'])
            df_sell_sum = df_sell_contracts.groupby(['CUSID', 'DUEPAYDT']).agg({'SETTTOT': 'sum'}).reset_index()
            df_sell_sum = strip_whitespace(df_sell_sum)
        else:
            df_sell_sum = pd.DataFrame(columns=['CUSID', 'DUEPAYDT', 'SETTTOT'])
        
        # 7. 處理議價交易資料
        if not df_buy_bargain.empty:
            # 按客戶、交割日、買賣別彙總
            df_buy_bargain_sum = df_buy_bargain.groupby(['CUSID', 'SETDAT', 'TXBS']).agg({'MTHAMT': 'sum'}).reset_index()
            
            # 調整金額：買進為正，賣出為負
            df_buy_bargain_sum['Adj_MTHAMT'] = np.where(
                df_buy_bargain_sum['TXBS'] == 'B',
                df_buy_bargain_sum['MTHAMT'],
                df_buy_bargain_sum['MTHAMT'] * -1
            )
            
            # 按客戶、交割日最終彙總
            df_buy_bargain_sum = df_buy_bargain_sum.groupby(['CUSID', 'SETDAT']).agg({'Adj_MTHAMT': 'sum'}).reset_index()
            df_buy_bargain_sum = strip_whitespace(df_buy_bargain_sum)
            print(df_buy_bargain_sum)
        else:
            df_buy_bargain_sum = pd.DataFrame(columns=['CUSID', 'SETDAT', 'Adj_MTHAMT'])
        
        return df_buy_sum, df_buy_bargain_sum, df_sell_sum, tday_plus_1, tday_plus_2
        
    except Exception as e:
        print(f"取得清算明細時發生錯誤: {e}")
        # 返回空的DataFrame
        empty_df = pd.DataFrame()
        return empty_df, empty_df, empty_df, "", ""
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass

def get_today_trade_detail(tday_str):
    conn = get_400_conn()
    df_today_trade_buy = pd.read_sql(f"""
        SELECT PRDID, CUSID, CBCODE, CBTQTY, STORQTY, TRDATE, SETDATE, OPTEXDT, PERRATE, CBTCOST, CBPER, CBTUPRM, PREMTOT, CBAMT FROM FSPFLIB.ASPROD WHERE TRDATE = '{tday_str}' AND TXTYPE = 'ASO'
    """, conn) # 新作契約
    df_today_trade_sell = pd.read_sql(f"""
        SELECT SEQNO, PRDID as PRDID_SELL, CUSID, CBCODE, DUEDATE, DUEPAYDT, CANTYPE, CANMODE, DEUQTY, CBTPDT, CBTPPRI, CANRATE, AVEPRICE, SETTTOT, DIFAMT FROM FSPFLIB.ASSURR WHERE DUEDATE = '{tday_str}'
    """, conn) # 提解契約

    prdids = df_today_trade_sell['PRDID_SELL'].unique().tolist()
    
    if prdids:
        prdids_str = "', '".join(str(x) for x in prdids)
        df_qty_left = pd.read_sql(
            f"SELECT PRDID as PRDID_QTY_LEFT, STORQTY as QTY_LEFT FROM FSPFLIB.ASPROD WHERE PRDID IN ('{prdids_str}')", conn
        ) # 剩餘庫存
    else:
        df_qty_left = pd.DataFrame(columns=['PRDID_QTY_LEFT', 'QTY_LEFT'])

    df_today_trade_sell = calculate_exercise_price(df_today_trade_sell) # 計算履約價

    df_today_trade = strip_whitespace(pd.concat([df_today_trade_buy, df_today_trade_sell])) # 合併新作和提解契約
    df_today_trade = df_today_trade.merge(df_qty_left, left_on='PRDID_SELL', right_on='PRDID_QTY_LEFT', how='left') # 合併剩餘庫存
    df_today_trade = df_today_trade.rename(columns={
        'PRDID': '新作契約編號',
        'CUSID': '客戶ID',
        'CBCODE': 'CB代號',
        'CBTQTY': '成交張數',
        'STORQTY': '庫存張數',
        'TRDATE': '交易日',
        'SETDATE': '交割日',
        'OPTEXDT': '選擇權到期日',
        'PERRATE': '履約利率',
        'CBTCOST': '成交均價',
        'CBPER': '百元價',
        'CBTUPRM': '單位權利金',
        'PREMTOT': '權利金總額',
        'SEQNO': '解約契約編號',
        'PRDID_SELL': '原單契約編號',
        'DUEDATE': '交易日_賣出',
        'DUEPAYDT': '交割日_賣出',
        'CBTPDT': '賣回日',
        'CBTPPRI': '賣回價',
        'CANTYPE': '解約類別', #0 = 到期未履約, 3 = 提前到期
        'CANMODE': '履約方式', #1 = 現金履約, 2 = 實物履約
        'DEUQTY': '履約張數',
        'CANRATE': '履約利率_賣出',
        'AVEPRICE': '成交均價_賣出',
        'SETTTOT': '交割金額',
        'DIFAMT': '履約損益',
        'QTY_LEFT': '剩餘張數',
    })

    df_today_trade['履約方式'] = np.where(
        df_today_trade['履約方式'] == '2', 
        '實物履約',
        np.where(df_today_trade['履約方式'] == '1',
            np.where(df_today_trade['解約類別'] == '0', '到期未履約', 
            np.where(df_today_trade['解約類別'] == '3', '提前到期', '現金結算')
            ),
            df_today_trade['履約方式']  # 其他情況保持原值
        )
    )
    
    df_today_trade['履約損益'] = np.where(df_today_trade['履約方式'] == '實物履約', 0, df_today_trade['履約損益'])

    df_today_bargain = pd.read_sql(f"""
        SELECT TXDATE, CUSID, ORDERNO, SETDAT, TXBS, STKID, MTHQTY, PRICE, MTHAMT FROM FSPFLIB.ASBARG WHERE TXDATE = '{tday_str}'
    """, conn)
    df_today_bargain = df_today_bargain.rename(columns={
        'TXDATE': '成交日',
        'CUSID': '客戶ID',
        'SETDAT': '交割日',
        'TXBS': '買/賣',
        'ORDERNO': '單據編號',
        'STKID': 'CB代號',
        'MTHQTY': '議價張數',
        'PRICE': '議價價格',
        'MTHAMT': '議價金額',
    })
    df_today_bargain = strip_whitespace(df_today_bargain)
    
    return df_today_trade, df_today_bargain

def calculate_exercise_price(df_today_trade_sell):
    # 將 YYYYMMDD 格式的字符串轉換為 datetime 對象
    df_today_trade_sell['CBTPDT'] = pd.to_datetime(df_today_trade_sell['CBTPDT'], format='%Y%m%d')
    df_today_trade_sell['DUEPAYDT'] = pd.to_datetime(df_today_trade_sell['DUEPAYDT'], format='%Y%m%d')
    
    # 計算履約價
    df_today_trade_sell['履約價'] = round(df_today_trade_sell['CBTPPRI'] - (df_today_trade_sell['CANRATE'] * (df_today_trade_sell['CBTPDT'] - df_today_trade_sell['DUEPAYDT'] + pd.Timedelta(days=1)).dt.days / 365), 2)
    return df_today_trade_sell

def check_each01():
    conn = get_400_conn()
    df_each01 = strip_whitespace(pd.read_sql("SELECT * FROM fspflib.FSPEACH01", conn))
    ad = df_each01['ADMARK'].fillna('').astype(str).str.strip()
    rc = df_each01['RCODE'].fillna('').astype(str).str.strip()

    # 依條件產生 ifbankok
    conditions = [
        (ad == 'A') & (rc.isin(['0', '4'])),
        (ad == 'A') & (rc != '0'),
        (ad == 'D') & (rc == '0'),
        (ad == 'D') & (rc == '')
    ]
    choices = [
        'Y',
        '已壓A，送件中',
        '已取消扣款授權成功',
        '已壓D，送件中(取消扣款授權)'
    ]

    df_each01['IFBANKOK'] = np.select(conditions, choices, default='未壓A')
    conn.close()
    return df_each01

def read_today_bargain_and_execute():
    tday_str = datetime.today().strftime("%Y%m%d")
    #tday_str = '20251204'
    df_today_bargain = pd.read_sql(f"""
        SELECT TXDATE, CUSID, ORDERNO, SETDAT, TXBS, STKID, MTHQTY, PRICE, MTHAMT FROM FSPFLIB.ASBARG WHERE TXDATE = '{tday_str}'
    """, get_400_conn())

    df_today_bargain = strip_whitespace(df_today_bargain)
    cuslist = df_today_bargain['CUSID'].unique().tolist()
    cus_info_all = strip_whitespace(get_customer_bank_and_email(cuslist))
    df_today_bargain = df_today_bargain.merge(cus_info_all[['CUSID', 'CUSNAME', 'BNKNAME', 'BNKBRH', 'BNKACTNO', 'CENTERNO']], left_on='CUSID', right_on='CUSID', how='left')

    df_cbname = strip_whitespace(pd.read_sql(f"""
        SELECT BDE010, BDE015 FROM FSPFLIB.ASBDEM
    """, get_400_conn()))

    df_today_bargain = df_today_bargain.merge(df_cbname, left_on='STKID', right_on='BDE010', how='left')

    df_today_bargain = df_today_bargain.rename(columns={
        'TXDATE': '成交日',
        'CUSID': '客戶ID',
        'ORDERNO': '單據編號',
        'SETDAT': '交割日',
        'TXBS': '統一證買進/賣出',
        'STKID': 'CB代號',
        'MTHQTY': '張數',
        'PRICE': '價格',
        'MTHAMT': '金額',
        'CUSNAME': '交易對手',
        'BNKNAME': '銀行',
        'BNKBRH': '分行',
        'BNKACTNO': '銀行帳號',
        'CENTERNO': '集保帳號', 
        'BDE015': 'CB名稱',
    })

    # 計算成交日到交割日之間的business day數，使用 next_business_day 作為依據
    def calc_t_plus(row, start_col, end_col):
        start_day = pd.to_datetime(row[start_col])
        end_day = pd.to_datetime(row[end_col])
        t_plus = 0
        while start_day < end_day:
            start_day = next_business_day(start_day, 1)
            t_plus += 1
        return f'T+{t_plus}'
    df_today_bargain['T+?'] = df_today_bargain.apply(lambda row: calc_t_plus(row, '成交日', '交割日'), axis=1)
    df_today_bargain['標的'] = df_today_bargain['CB代號'] + ' ' + df_today_bargain['CB名稱']
    df_today_bargain['銀行帳號'] = df_today_bargain['銀行'] + df_today_bargain['分行'] + df_today_bargain['銀行帳號']
    
    # 格式化：標的、張數、金額去掉.0，金額添加千分位符
    df_today_bargain['標的'] = df_today_bargain['標的'].astype(str).str.replace(r'\.0\b', '', regex=True)
    df_today_bargain['張數'] = df_today_bargain['張數'].astype(str).str.replace(r'\.0\b', '', regex=True)
    # 金額：先轉為數值，去掉.0，再添加千分位符
    df_today_bargain['金額'] = pd.to_numeric(df_today_bargain['金額'], errors='coerce').fillna(0)
    df_today_bargain['金額'] = df_today_bargain['金額'].apply(lambda x: f"{int(x):,}" if x == int(x) else f"{x:,.2f}")

    df_today_bargain = df_today_bargain[['客戶ID', '統一證買進/賣出', '交易對手', '單據編號', '集保帳號', '標的', '張數', '價格', '金額', '銀行帳號', '交割日', 'T+?']]

#==================================實物履約=================================
    df_today_execute = strip_whitespace(pd.read_sql(f"""
        SELECT SEQNO, PRDID, CUSID, CBCODE, DUEDATE, DUEPAYDT, PERPRICE, DEUQTY, SETTTOT FROM FSPFLIB.ASSURR WHERE DUEDATE = '{tday_str}' AND CANMODE = '2'
    """, get_400_conn()))
    cuslist = df_today_execute['CUSID'].unique().tolist()
    cus_info_all = strip_whitespace(get_customer_bank_and_email(cuslist))
    df_today_execute = df_today_execute.merge(cus_info_all[['CUSID', 'CUSNAME']], left_on='CUSID', right_on='CUSID', how='left')
    df_today_execute = df_today_execute.merge(df_cbname, left_on='CBCODE', right_on='BDE010', how='left')
    df_today_execute = df_today_execute.rename(columns={
        'SEQNO': '解約契約編號',
        'PRDID': '原單契約編號',
        'CUSID': '客戶ID',
        'CBCODE': 'CB代號',
        'DUEDATE': '履約日',
        'DUEPAYDT': '履約交割日',
        'PERPRICE': '履約價',
        'DEUQTY': '履約張數',
        'SETTTOT': '買賣成交金額',
        'CUSNAME': '客戶名稱',
        'BDE015': 'CB名稱',
    })

    df_today_execute['標的'] = df_today_execute['CB代號'] + ' ' + df_today_execute['CB名稱']
    df_today_execute['T+?'] = df_today_execute.apply(lambda row: calc_t_plus(row, '履約日', '履約交割日'), axis=1)
    df_today_execute['履約張數'] = df_today_execute['履約張數'].astype(str).str.replace(r'\.0\b', '', regex=True)
    df_today_execute['買賣成交金額'] = pd.to_numeric(df_today_execute['買賣成交金額'], errors='coerce').fillna(0)
    df_today_execute['買賣成交金額'] = df_today_execute['買賣成交金額'].apply(lambda x: f"{int(x):,}" if x == int(x) else f"{x:,.2f}")

    df_today_execute = df_today_execute[['客戶名稱', '解約契約編號', '原單契約編號', '履約日', '履約交割日', '標的', '履約張數', '履約價', '買賣成交金額', 'T+?']]

    return df_today_bargain, df_today_execute
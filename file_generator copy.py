import pandas as pd
import pyodbc
from format_utils import strip_whitespace
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
            "SELECT * FROM FSPFLIB.ASPROD WHERE CUSID = 'S123581497'", 
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
    
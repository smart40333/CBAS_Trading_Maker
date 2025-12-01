import pandas as pd
import numpy as np
from datetime import datetime, date
import calendar
from decimal import Decimal
import pyodbc

def strip_whitespace(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    return df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

def format_date(date_str: str) -> str:
    if not date_str or len(str(date_str)) != 8:
        return date_str
    s = str(date_str)
    return f"{s[:4]}/{s[4:6]}/{s[6:8]}"


def strip_trailing_zeros(val):
    """將數字轉為不帶多餘 0 的字串，並避免浮點誤差顯示如 26.380000000000003"""
    try:
        if pd.isna(val):
            return ""
        # 整數型態
        if isinstance(val, (int, np.integer)):
            return str(int(val))
        # 浮點型態：用固定小數格式轉字串再去除尾零，避免二進位浮點誤差
        if isinstance(val, (float, np.floating)):
            s = f"{float(val):.12f}"  # 足夠精度後去尾零
            return s.rstrip('0').rstrip('.')
        # 字串：嘗試用 Decimal 精確處理，再格式化
        if isinstance(val, str):
            try:
                d = Decimal(val)
                # normalize 轉為最簡表示，再以 'f' 格式避免科學記號
                s = format(d.normalize(), 'f')
                return s.rstrip('0').rstrip('.') if '.' in s else s
            except Exception:
                return val
        return str(val)
    except (ValueError, TypeError):
        return str(val)

def convert_to_chinese_amount(amount) -> str:
    chinese_numbers = ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖']

    def convert_section(num: int) -> str:
        if num == 0:
            return ''
        result = ''
        str_num = str(num).zfill(4)
        need_zero = False
        if str_num[0] != '0':
            result += chinese_numbers[int(str_num[0])] + '仟'
            need_zero = True
        if str_num[1] != '0':
            if need_zero and str_num[0] == '0':
                result += '零'
            result += chinese_numbers[int(str_num[1])] + '佰'
            need_zero = True
        elif need_zero and (str_num[2] != '0' or str_num[3] != '0'):
            result += '零'
            need_zero = False
        if str_num[2] != '0':
            if need_zero and str_num[1] == '0':
                result += '零'
            result += chinese_numbers[int(str_num[2])] + '拾'
            need_zero = True
        elif need_zero and str_num[3] != '0':
            result += '零'
            need_zero = False
        if str_num[3] != '0':
            if need_zero and str_num[2] == '0':
                result += '零'
            result += chinese_numbers[int(str_num[3])]
        return result

    amount_str = str(amount).replace(',', '').replace(' ', '')
    try:
        amount_int = int(float(amount_str))
        if amount_int == 0:
            return '零元整'
        yi = amount_int // 100000000
        wan = (amount_int % 100000000) // 10000
        ge = amount_int % 10000
        result = ''
        if yi > 0:
            result += convert_section(yi) + '億'
        if wan > 0:
            if yi > 0 and wan < 1000:
                result += '零'
            result += convert_section(wan) + '萬'
        if ge > 0:
            if result and ge < 1000:
                result += '零'
            result += convert_section(ge)
        if not result:
            result = '零'
        while '零零' in result:
            result = result.replace('零零', '零')
        if result.endswith('零'):
            result = result[:-1]
        result += '元整'
        return result
    except ValueError:
        return '金額格式錯誤'


def float_to_str_maxlen(val, maxlen=11):
    """將浮點數轉成長度不超過 maxlen 的字串（含小數點）"""
    try:
        f = float(val)
        s_int = str(int(abs(f)))
        int_len = len(s_int)
        sign = '-' if f < 0 else ''
        dot = 1 if f % 1 != 0 else 0
        decimals = maxlen - len(sign) - int_len - dot
        if decimals < 0:
            return sign + s_int[:maxlen - len(sign)]
        fmt = f"{{0:.{decimals}f}}"
        s = fmt.format(f)
        if '.' in s:
            s = s.rstrip('0').rstrip('.')
        if len(s) > maxlen:
            s = s[:maxlen]
        return s
    except:
        return str(val)


def read_holiday_list():
    connSQL = pyodbc.connect(driver='ODBC Driver 18 for SQL Server', server='10.72.228.139', user='sa', password='Self@pscnet', database='CBAS', TrustServerCertificate='yes')
    df_holiday = pd.read_sql("SELECT * FROM HolidayList", connSQL)
    connSQL.close()
    return df_holiday

def next_business_day(start_date: datetime, days: int) -> datetime:
    """計算指定日期後的第N個工作日"""
    df_holiday = read_holiday_list()
    holiday_set = set()
    for date_str in df_holiday['Date']:
        if isinstance(date_str, str):
            # 支援 2024/6/10 或 2024-06-10
            date_str = date_str.replace('-', '/')
            parts = date_str.split('/')
            if len(parts) == 3:
                y, m, d = map(int, parts)
                holiday_set.add(datetime(y, m, d).date())
    
    current = start_date
    added = 0
    while added < days:
        current += pd.Timedelta(days=1)
        if current.weekday() < 5 and current.date() not in holiday_set:
            added += 1
    return current 

def prev_business_day(start_date: datetime, days: int) -> datetime:
    """計算指定日期前的第N個工作日"""
    df_holiday = read_holiday_list()
    holiday_set = set()
    for date_str in df_holiday['Date']:
        if isinstance(date_str, str):
            date_str = date_str.replace('-', '/')
            parts = date_str.split('/')
            if len(parts) == 3:
                y, m, d = map(int, parts)
                holiday_set.add(datetime(y, m, d).date())
    current = start_date
    deducted = 0
    while deducted < days:
        current -= pd.Timedelta(days=1)
        if current.weekday() < 5 and current.date() not in holiday_set:
            deducted += 1
    return current


def calculate_expired_exercise_price(sellback_price: float, year_period: float, exercise_rate: float) -> float:
    """計算到期履約價"""
    return round(sellback_price - (100 * year_period * exercise_rate / 100), 2)


def calculate_expired_year_period(sellback_date: str, settlement_date: str) -> float:
    """計算到期年期 = (賣回日 - 交割日) / 365"""
    try:
        sellback_dt = pd.to_datetime(sellback_date, format='%Y%m%d')
        settlement_dt = pd.to_datetime(settlement_date, format='%Y%m%d')
        year_period = np.maximum((sellback_dt - settlement_dt).days + 1, 0) / 365
        return year_period
    except Exception as e:
        print(f"計算到期年期時發生錯誤: {e}")
        return 0.0


def format_expired_contract_data(df_expired: pd.DataFrame, target_date: str, df_quote: pd.DataFrame) -> pd.DataFrame:
    """格式化到期契約資料"""
    try:
        # 重新命名和選擇欄位
        df_final = df_expired.copy()
        df_final = df_final.rename(columns={
            'PRDID': '原單契約編號',
            'CUSID': '客戶ID',
            'CUSNAME': '客戶名稱',
            'STORQTY': '原庫存張數',
            'CBTPDT': '賣回日',
            'CBTPPRI': '賣回價',
            'PERRATE': '履約利率',
            'OPTEXDT': '選擇權到期日',
        })

        # 設定固定欄位
        tday = datetime.now().strftime('%Y%m%d')
        df_final['交易日期'] = tday
        tday_datetime = pd.to_datetime(tday, format='%Y%m%d')
        settle_date = next_business_day(tday_datetime, 2)
        df_final['交割日期'] = settle_date.strftime('%Y%m%d')
        df_final['解約類別'] = '0'
        df_final['履約方式'] = '1'
        df_final['提前履約賠償金'] = '0'

        # 計算年期和履約價
        df_final['年期'] = df_final['賣回日'].apply(
            lambda x: calculate_expired_year_period(x, settle_date.strftime('%Y%m%d'))
        )
        df_final['履約價'] = df_final.apply(
            lambda row: calculate_expired_exercise_price(
                float(row['賣回價']), row['年期'], float(row['履約利率'])
            ), axis=1
        )

        df_final['成交均價'] = df_final['履約價']
        df_final['選擇權交割單價'] = df_final['成交均價'] - df_final['履約價']
        
        # 確保剩餘到期張數存在
        if '剩餘到期張數' in df_final.columns:
            df_final['交割總金額'] = df_final['剩餘到期張數'] * df_final['選擇權交割單價']
        else:
            df_final['交割總金額'] = 0
            
        df_final['錄音時間'] = ''

        return df_final

    except Exception as e:
        print(f"格式化到期契約資料時發生錯誤: {e}")
        return pd.DataFrame() 
    
def edate(start_date: date, months: int) -> date:
    # 計算新年月
    month = start_date.month - 1 + months
    year = start_date.year + month // 12
    month = month % 12 + 1
    
    # 該月的最後一天
    last_day = calendar.monthrange(year, month)[1]
    
    # 如果原日期比該月的最後一天還大，就用該月最後一天
    day = min(start_date.day, last_day)
    
    return date(year, month, day)

def cusid_to_padded(cusid: str) -> str:
    s = "" if cusid is None else str(cusid)
    return s.ljust(12)  # 右側補空白到長度 12

def format_number_to_11(value, max_length=11):
    if pd.isna(value):
        return value
    
    float_val = float(value)
    
    # 如果是整数，直接返回
    if float_val == int(float_val):
        return str(int(float_val))
    
    # 计算整数部分长度
    int_part = str(int(float_val))
    int_length = len(int_part)
    
    # 计算可以保留的小数位数
    max_decimals = max_length - int_length - 1  # -1 for decimal point
    
    if max_decimals <= 0:
        # 整数部分太长，返回整数
        return str(int(round(float_val)))
    
    # 尝试从最大可能的小数位数开始递减
    for decimals in range(min(max_decimals, 9), -1, -1):
        formatted = f"{float_val:.{decimals}f}".rstrip('0').rstrip('.')
        if len(formatted) <= max_length:
            return formatted
    
    return str(int(round(float_val)))


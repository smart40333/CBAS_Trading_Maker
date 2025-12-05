import subprocess
import sys
import traceback
import pandas as pd
from datetime import datetime, timedelta
from PyQt5.QtWidgets import QMessageBox
import win32com.client as win32
# 導入必要的模組
from db_access import get_customer_bank_and_email, get_clearing_detail, get_today_trade_detail, check_each01, read_today_bargain_and_execute, get_400_conn
from file_reader import load_quote
from file_generator import generate_trade_notice_template
from format_utils import cusid_to_padded, strip_whitespace
from envs import trade_notice_dir
import numpy as np

def send_email(body, subject, to, attpath=None, html_body=None):
    """
    寄信
    
    Args:
        body: 郵件正文（純文字）
        subject: 郵件主旨
        to: 收件人（可以是單個郵件地址或分號分隔的多個地址）
        attpath: 附件路徑，可以是：
            - None: 無附件
            - str: 單個附件路徑
            - list: 多個附件路徑的列表
        html_body: HTML格式的郵件正文（如果提供，會優先使用此而非body）
    """
    outlook = win32.Dispatch('outlook.application')
    account = None
    # 尋找 PSC.CBAS@uni-psg.com 這個帳號
    for acc in outlook.Session.Accounts:
        if acc.SmtpAddress.lower() == 'psc.cbas@uni-psg.com':
            account = acc
            break
    if account is None:
        raise Exception("Outlook找不到PSC.CBAS@uni-psg.com帳號，請確認Outlook已登入該帳號")
    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # 64209 = 'SendUsingAccount'
    mail.Subject = subject
    if html_body:
        mail.HTMLBody = html_body
    else:
        mail.Body = body
    mail.To = to
    #mail.To = 'wayne.huang@uni-psg.com'
    
    # 處理附件：支援單個附件或多個附件
    if attpath:
        if isinstance(attpath, str):
            # 單個附件
            mail.Attachments.Add(attpath)
        elif isinstance(attpath, list):
            # 多個附件
            for path in attpath:
                if path:  # 確保路徑不為空
                    mail.Attachments.Add(path)
    
    mail.Send()

def send_today_trade_email(output_text_edit, parent):
    """寄信: 今日交易筆數"""
    try:
        tday_str = datetime.today().strftime("%Y%m%d")
        df_today_trade, df_today_bargain = get_today_trade_detail(tday_str)
        #print(df_today_trade)
        buy_qty = len(df_today_trade[df_today_trade['新作契約編號'].notna()])
        df_sell_prdids = df_today_trade[(df_today_trade['解約契約編號'].notna()) & (df_today_trade['履約方式'] == '現金結算')]['原單契約編號'].unique()
        
        # 查詢 ASPROD 表，找出這些 PRDID 且 TXTYPE == 'ASO' 的記錄
        if len(df_sell_prdids) > 0:
            prdids_list = "','".join(df_sell_prdids)
            conn = get_400_conn()
            df_sell_asprod = strip_whitespace(pd.read_sql(
                f"SELECT * FROM FSPFLIB.ASPROD WHERE PRDID IN ('{prdids_list}') AND TXTYPE = 'ASO'",
                conn
            ))
            conn.close()
        else:
            df_sell_asprod = pd.DataFrame()
        
        sell_qty = len(df_sell_asprod)
        exe_qty = len(df_today_trade[(df_today_trade['解約契約編號'] != '') & (df_today_trade['解約契約編號'].notna()) & (df_today_trade['履約方式'] == '實物履約')])
        #print(buy_qty, sell_qty, exe_qty)
        body = f"""
        Dear All, 

        已上傳{buy_qty}筆新作交易；{sell_qty}筆現金提解交易；{exe_qty}筆實物履約。(交割日期、金額另一封信補充)。(請留意可能有扣款)； 
        
        以上，再麻煩協助後續作業，謝謝。 
        """
        maillist = "MIKE@uni-psg.com;IRENELIN@uni-psg.com;10176@uni-psg.com;AMMYCHANG@uni-psg.com;MEILAN@uni-psg.com;CATHERINE@uni-psg.com;GRACE.ROSA@uni-psg.com;NBDCHANG@uni-psg.com;WAYNE.HUANG@uni-psg.com;CHARLESP@uni-psg.com;PANGYEN@uni-psg.com;LINDY00@uni-psg.com;EMMA@uni-psg.com;YUNA.WU@uni-psg.com;CHUN-HUEI@uni-psg.com;P5480@uni-psg.com;95105@uni-psg.com;CHANTAL.CHU@uni-psg.com;YITAN9593@uni-psg.com;12267@uni-psg.com;VANASSA@uni-psg.com;KMJUI.TSAI@uni-psg.com"
        tday_str = datetime.today().strftime("%Y%m%d")
        send_email(body, f"CBAS本日交易{tday_str}", maillist)
        output_text_edit.append(f"CBAS本日交易{tday_str}已寄出")
        QMessageBox.information(parent, "成功", f"CBAS本日交易{tday_str}已寄出")
        
        # 發送第一封郵件成功後，再發送上傳檔郵件
        send_upload_file_email(output_text_edit, parent)
        
    except Exception as e:
        output_text_edit.append(f"✗ 執行寄信程式時發生錯誤：{e}\n")
        QMessageBox.critical(parent, "錯誤", f"執行寄信程式時發生錯誤：{e}")
        # 即使第一封郵件失敗，也嘗試發送上傳檔郵件（因為上傳檔是獨立的）
        try:
            send_upload_file_email(output_text_edit, parent)
        except Exception as e2:
            output_text_edit.append(f"✗ 發送上傳檔郵件時發生錯誤：{e2}\n")

def send_upload_file_email(output_text_edit, parent):
    """寄信: 上傳檔"""
    try:
        tday_str = datetime.today().strftime("%Y%m%d")
        to = 'NBDCHANG@uni-psg.com;RIVER@uni-psg.com;14122@uni-psg.com'
        subject = f"CBAS新作提解上傳檔{tday_str}"
        body = f"CBAS新作提解上傳檔{tday_str}已寄出"
        # 多個附件使用列表
        attpath = [
            rf'\\10.72.228.112\cbas業務公用區\CBAS上傳檔\新作上傳檔.csv',
            rf'\\10.72.228.112\cbas業務公用區\CBAS上傳檔\解約上傳檔.csv'
        ]
        send_email(body, subject, to, attpath=attpath, html_body=None)
        output_text_edit.append(f"✓ CBAS新作提解上傳檔{tday_str}已寄出")
        QMessageBox.information(parent, "成功", f"CBAS新作提解上傳檔{tday_str}已寄出")
    except Exception as e:
        output_text_edit.append(f"✗ 執行寄信程式時發生錯誤：{e}\n")

def send_bargain_trade_email(output_text_edit, parent):
    """寄信: 議價交易"""
    try:
        tday_str = datetime.today().strftime("%Y%m%d")
        df_today_bargain, df_today_execute = read_today_bargain_and_execute()
        bargain_qty_t0 = len(df_today_bargain[df_today_bargain['T+?'] == 'T+0'])
        bargain_qty_t1 = len(df_today_bargain[df_today_bargain['T+?'] == 'T+1'])
        execute_qty_t1 = len(df_today_execute[df_today_execute['T+?'] == 'T+1'])

        df_compare = pd.read_excel(
            r'\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\議價檔ASBARG上傳檔.xlsx',
            dtype=str
        )
        
        # 在 df_compare 中創建組合 key：CUSID + STOCK + MTHQTY
        df_compare['match_key'] = df_compare['CUSID'].astype(str).str.strip() + '|' + df_compare['STOCK'].astype(str).str.strip() + '|' + df_compare['MTHQTY'].astype(str).str.strip()
        df_compare['集保帳號_比對'] = df_compare['BRKID'].astype(str).str.strip() + df_compare['ACCTNO'].astype(str).str.strip()
        
        # 在 df_today_bargain 中創建組合 key：客戶ID + CB代號 + 張數
        # 根據標的前幾碼是否為數字來決定切片位置
        def get_cb_code_prefix(value):
            """根據前幾碼是否為數字決定切片位置：直接取到第幾位是數字就是取到第幾位"""
            value_str = str(value).strip()
            # 從第一位開始檢查連續有多少位是數字
            digit_count = 0
            for char in value_str:
                if char.isdigit():
                    digit_count += 1
                else:
                    break
            
            # 如果找到連續的數字，就取到那個位置
            if digit_count > 0:
                return value_str[:digit_count]
            # 如果沒有數字，預設使用前4碼
            return value_str[:4] if len(value_str) >= 4 else value_str
        
        df_today_bargain['標的_前綴'] = df_today_bargain['標的'].apply(get_cb_code_prefix)
        df_today_bargain['match_key'] = df_today_bargain['客戶ID'].astype(str).str.strip() + '|' + df_today_bargain['標的_前綴'].astype(str).str.strip() + '|' + df_today_bargain['張數'].astype(str).str.strip()
        
        # 使用組合 key 進行 merge
        df_today_bargain = pd.merge(
            df_today_bargain, 
            df_compare[['match_key', '集保帳號_比對']], 
            on='match_key', 
            how='left'
        )
        
        # 如果集保帳號不一樣，使用 df_compare 的集保帳號
        df_today_bargain['集保帳號'] = np.where(
            (df_today_bargain['集保帳號_比對'].notna()) & 
            (df_today_bargain['集保帳號_比對'] != df_today_bargain['集保帳號']),
            df_today_bargain['集保帳號_比對'],
            df_today_bargain['集保帳號']
        )
        
        # 清理臨時欄位
        df_today_bargain.drop(columns=['match_key', '集保帳號_比對', '標的_前綴'], inplace=True)
        print(df_today_bargain)
        df_today_bargain_html = df_today_bargain.to_html(index=False, classes='table', table_id='bargain_table')
        df_today_execute_html = df_today_execute.to_html(index=False, classes='table', table_id='execute_table')
        html_body = f"""
        <p>Dear All,</p>
        
        <p><b>議價交易</b></p>
        {df_today_bargain_html}
        
        <p><b>實物履約</b></p>
        {df_today_execute_html}
        """
        body = "Dear All,\n\n請查看HTML格式的郵件內容。"
        maillist = 'vanassa@uni-psg.com; 12267@uni-psg.com; MIKE@uni-psg.com; DANIEL02@uni-psg.com; YUNA.WU@uni-psg.com; IRENELIN@uni-psg.com; 10176@uni-psg.com; CHUN-HUEI@uni-psg.com; AMMYCHANG@uni-psg.com; MEILAN@uni-psg.com; ERICCHEN@uni-psg.com; CATHERINE@uni-psg.com; GRACE.ROSA@uni-psg.com; XX24923051@uni-psg.com; NBDCHANG@uni-psg.com; IRENEHUANG@uni-psg.com; VANASSA@uni-psg.com; CHARLESP@uni-psg.com; PANGYEN@uni-psg.com; YIHUI@uni-psg.com; YUCHIN.HSUEH@uni-psg.com; LINDY00@uni-psg.com; P5480@uni-psg.com; EMMA@uni-psg.com; 95105@uni-psg.com; CHANTAL.CHU@uni-psg.com; YITAN9593@uni-psg.com; YICIH@uni-psg.com; KMJUI.TSAI@uni-psg.com'
        if bargain_qty_t0 > 0 or bargain_qty_t1 > 0 or execute_qty_t1 > 0:
            attpath = rf'\\10.72.228.112\cbas業務公用區\!!!交易作業區!!!\議價交易\議價交易內部通知\議價交易_{tday_str}.pdf'
            #print(df_today_bargain_html)
            #send_email(body, f"{tday_str}__議價交易{bargain_qty_t0}筆T+0，{bargain_qty_t1}筆T+1，實物履約{execute_qty_t1}筆T+1_附件", maillist, attpath, html_body)
        else:
            send_email(body, f"{tday_str}__無議價交易及實物履約", maillist, attpath=None, html_body=html_body)

        output_text_edit.append(f"✓ 議價交易郵件已寄出：{bargain_qty_t0}筆T+0，{bargain_qty_t1}筆T+1，實物履約{execute_qty_t1}筆T+1\n")
        QMessageBox.information(parent, "成功", f"議價交易郵件已寄出：{bargain_qty_t0}筆T+0，{bargain_qty_t1}筆T+1，實物履約{execute_qty_t1}筆T+1")

    except Exception as e:
        output_text_edit.append(f"✗ 執行寄信程式時發生錯誤：{e}\n")
        QMessageBox.critical(parent, "錯誤", f"執行寄信程式時發生錯誤：{e}")

def generate_today_detail(output_text_edit, parent):
    """產檔: 產今日成交明細（依客戶ID分檔）"""
    try:
        # 由表格取出 DataFrame
        tday = datetime.today()
        #tday = tday - timedelta(days=1)
        tday_str = tday.strftime("%Y%m%d")
        #tday_str = '20251201'
        df_today_trade, df_today_bargain = get_today_trade_detail(tday_str)

        cusid_list = df_today_trade['客戶ID'].unique()
        cus_info_all = get_customer_bank_and_email(cusid_list)

        df_quote, duplicate_cb, df_cbinfo = load_quote()
        df_today_trade = df_today_trade.merge(cus_info_all[['CUSID', 'CUSNAME']], left_on='客戶ID', right_on='CUSID', how='left')
        df_today_trade = df_today_trade.merge(df_quote[['CB代號', 'CB名稱']], left_on='CB代號', right_on='CB代號', how='left')

        df_buy_sum, df_buy_bargain_sum, df_sell_sum, tday_plus_1, tday_plus_2 = get_clearing_detail(tday)
        df_today_clearing_money = pd.concat([df_buy_sum, df_buy_bargain_sum, df_sell_sum])
        df_today_bargain = df_today_bargain.merge(df_quote[['CB代號', 'CB名稱']], left_on='CB代號', right_on='CB代號', how='left')
        df_each01 = check_each01()
        
        # 用於累積所有客戶的交割資訊
        all_clearing_info_list = []
        
        for who in cusid_list:
            df_today_trade_person = df_today_trade[df_today_trade['客戶ID'] == who]
            df_today_bargain_person = df_today_bargain[df_today_bargain['客戶ID'] == who]
            df_today_clearing_money_person = df_today_clearing_money[df_today_clearing_money['CUSID'] == who]
            cus_info_person = cus_info_all[cus_info_all['CUSID'] == who]
            
            # 檢查 df_each01 中是否有該客戶的資料
            df_each01_filtered = df_each01[df_each01['CUSID'] == who]
            if len(df_each01_filtered) > 0:
                ifbankok_person = df_each01_filtered['IFBANKOK'].values[0]
            else:
                # 如果找不到該客戶，使用預設值 'N'
                ifbankok_person = 'N'
            
            excel_path, pdf_path, df_clearing_info = generate_trade_notice_template(
                cus_id=who,
                cus_info=cus_info_person,
                df_today_trade=df_today_trade_person,
                df_today_clearing_money=df_today_clearing_money_person,
                tday_str=tday_str,
                df_today_bargain=df_today_bargain_person,
                tday_plus_1=tday_plus_1,
                tday_plus_2=tday_plus_2,
                ifbankok=ifbankok_person,
            )
            
            # 累積交割資訊
            if not df_clearing_info.empty:
                all_clearing_info_list.append(df_clearing_info)
        
        #if excel_path:
        #    output_text_edit.append(f"✓ Excel檔產出完成：{excel_path}\n")
            if pdf_path:
                output_text_edit.append(f"✓ PDF檔產出完成：{pdf_path}\n")

        # 所有客戶處理完後，統一保存交割資訊
        if all_clearing_info_list:
            df_all_clearing_info = pd.concat(all_clearing_info_list, ignore_index=True)
            clearing_info_path = r'\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\交割資訊.xlsx'
            df_all_clearing_info.to_excel(clearing_info_path, index=False)
            output_text_edit.append(f"✓ 交割資訊已統一保存：{clearing_info_path}\n")
        
        QMessageBox.information(parent, "提示", "今日成交明細產檔完成！開始寄信")
        
        df_customer_money = pd.read_excel(r'\\10.72.228.112\cbas業務公用區\CBAS_Trading_Maker\交割資訊.xlsx')
        df_customer_money = df_customer_money[['客戶ID', '交易日期', '交割總額', '收付', '交割方式']]

        for who in cusid_list:
            df_customer_money_person = df_customer_money[df_customer_money['客戶ID'] == who]
            df_customer_money_person.drop(columns=['客戶ID'], inplace=True)
            cus_info_person = cus_info_all[cus_info_all['CUSID'] == who]
            to = cus_info_person['EMAIL'].values[0]
            subject = f"統一證券CBAS_當日成交明細_{tday_str}"
            
            # 生成表格 HTML
            table_html = df_customer_money_person.to_html(index=False, classes='table', table_id='customer_money_table', escape=False)
            
            # 構建完整的 HTML 內容，包含 CSS 樣式
            html_body = f"""
            <html>
            <head>
                <style>
                    body {{
                        font-family: "Microsoft JhengHei", "微軟正黑體", Arial, sans-serif;
                        font-size: 12pt;
                        line-height: 1.0;
                    }}
                    #customer_money_table {{
                        border-collapse: collapse;
                        table-layout: fixed;
                        width: 440px;           /* 寬度固定 px */
                        font-family: "Microsoft JhengHei", "微軟正黑體", Arial, sans-serif;
                        font-size: 12pt;
                        margin: 8px 0;
                    }}
                    #customer_money_table th {{
                        background-color: #E6F2FF;
                        color: #000000;
                        font-weight: bold;
                        text-align: center;
                        border: 1px solid #000000;
                        
                        /* --- 標題列高度設定 --- */
                        height: 20px;           /* 設定你要的高度 */
                        line-height: 20px;      /* 關鍵：行高設為跟高度一樣，文字會自動垂直置中 */
                        padding: 0px;           /* 關鍵：內距歸零 */
                        overflow: hidden;       /* 超出範圍隱藏 */
                        white-space: nowrap;    /* 強制不換行 (避免標題兩行撐開高度) */
                    }}
                    #customer_money_table td {{
                        text-align: center;
                        border: 1px solid #000000;
                        background-color: #FFFFFF;
                        
                        /* --- 資料列高度設定 --- */
                        height: 20px;           /* 設定你要的高度 */
                        line-height: 20px;      /* 關鍵：行高設為跟高度一樣 */
                        padding: 0px;           /* 關鍵：內距歸零 */
                        overflow: hidden;       /* 超出範圍隱藏 */
                        white-space: nowrap;    /* 強制不換行 (避免內容兩行撐開高度) */
                    }}
                    #customer_money_table tr:nth-child(even) td {{
                        background-color: #F9F9F9;
                    }}
                    #customer_money_table tr:hover td {{
                        background-color: #F0F8FF;
                    }}
                    
                    /* 寬度設定 (全部用 px) */
                    #customer_money_table th:nth-child(1), #customer_money_table td:nth-child(1) {{ width: 100px; }}
                    #customer_money_table th:nth-child(2), #customer_money_table td:nth-child(2) {{ width: 180px; }}
                    #customer_money_table th:nth-child(3), #customer_money_table td:nth-child(3) {{ width: 80px; }}
                    #customer_money_table th:nth-child(4), #customer_money_table td:nth-child(4) {{ width: 80px; }}
                </style>
            </head>
            <body>
                <p>貴客戶您好,</p>
                <p>當日成交部位資訊請參考附件，謝謝。</p>
                <p>以下為您近期的交割資訊：</p>
                {table_html}
                <p>此信件由系統自動發出，請勿直接回覆，如有任何疑問請聯絡相關業務人員。</p>
                <p>Regards,</p>
                <p>統一證券<br>
                黃暐庭02-27463670<br>
                周琬瑜02-27463727<br>
                Email PSC.CBAS@uni-psg.com<br>
                地址105台北市松山區東興路8號2樓</p>
            </body>
            </html>
            """
            cusname_full = cus_info_person['CUSNAME'].values[0][0] + 'Ｏ' + cus_info_person['CUSNAME'].values[0][-1]
            cusEmail = cus_info_person['EMAIL'].values[0][:4]
            if cusname_full in ['劉Ｏ漢', '劉Ｏ餘', '陳Ｏ美']:
                attpath = rf'{trade_notice_dir}\{tday_str}\EXCEL\{tday_str}_統一證券CBAS成交明細_{cusname_full}_{cusEmail}.xlsx'
            else:
                attpath = rf'{trade_notice_dir}\{tday_str}\PDF\{tday_str}_統一證券CBAS成交明細_{cusname_full}_{cusEmail}.pdf'
            print("寄出信件: ", cusname_full, cusEmail, subject, to, attpath)
            send_email(body=None, subject=subject, to=to, attpath=attpath, html_body=html_body)
            

        
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        tb = traceback.extract_tb(exc_tb)
        if tb:
            lineno = tb[-1].lineno
            output_text_edit.append(f"✗ 產檔過程發生錯誤：{e}（第 {lineno} 行）\n")
        else:
            output_text_edit.append(f"✗ 產檔過程發生錯誤：{e}\n")
        QMessageBox.critical(parent, "錯誤", f"產檔過程發生錯誤：{e}")

def generate_trade_confirmation(output_text_edit, parent):
    """產檔: 產交易確認書"""
    try:
        # 開啟Excel檔案並執行巨集
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        output_text_edit.append("=== 產交易確認書 ===\n")
        output_text_edit.append("正在開啟Excel檔案...\n")
        
        # 開啟Excel檔案（請根據實際路徑修改）
        workbook = excel.Workbooks.Open(r"\\10.72.228.112\cbas業務公用區\各式交易確認書範本.xlsm")
        
        output_text_edit.append("正在執行巨集 CallAll...\n")
        # 執行巨集
        excel.Run("CallAll")
        
        # 儲存並關閉
        workbook.Save()
        workbook.Close()
        excel.Quit()
        
        output_text_edit.append("✓ 交易確認書產檔完成！\n")
        QMessageBox.information(parent, "成功", "交易確認書產檔完成！")
    except Exception as e:
        output_text_edit.append(f"✗ 產檔過程發生錯誤：{e}\n")
        QMessageBox.critical(parent, "錯誤", f"產檔過程發生錯誤：{e}")

def send_control_table_email(output_text_edit, parent):
    """寄信: 控管表"""
    try:
        # 執行對應的.bat檔案
        result = subprocess.run([r'\\10.72.228.120\Py_Project\Bat\controlTableFinal.bat'], 
                                capture_output=True, text=True)
        
        if result.stdout:
            output_text_edit.append("=== 控管表寄信執行結果 ===\n")
            output_text_edit.append(f"標準輸出:\n{result.stdout}\n")
        QMessageBox.information(parent, "成功", "控管表寄信完成！")
    except Exception as e:
        QMessageBox.critical(parent, "錯誤", f"執行寄信程式時發生錯誤：{e}")

def send_customer_detail_email(output_text_edit, parent):
    """寄信: 客戶當日成交明細"""
    try:
        # 執行對應的.py檔案
        result = subprocess.run([r'\\10.72.228.120\Py_Project\Bat\DailyTradingDetail_1500.bat'], 
                                capture_output=True, text=True)
        if result.stdout:
            output_text_edit.append("=== 客戶當日成交明細寄信執行結果 ===\n")
            output_text_edit.append(f"標準輸出:\n{result.stdout}\n")
        QMessageBox.information(parent, "成功", "客戶當日成交明細寄信完成！")
    except Exception as e:
        output_text_edit.append(f"✗ 執行寄信程式時發生錯誤：{e}\n")
        QMessageBox.critical(parent, "錯誤", f"執行寄信程式時發生錯誤：{e}")

def send_customer_positions_email(output_text_edit, parent):
    """寄信: 客戶部位表"""
    try:
        # 執行對應的.py檔案
        result = subprocess.run([r'\\10.72.228.120\Py_Project\Bat\Gen_Cus_Positions.bat'], 
                                capture_output=True, text=True)
        if result.stdout:
            output_text_edit.append("=== 客戶部位表寄信執行結果 ===\n")
            output_text_edit.append(f"標準輸出:\n{result.stdout}\n")
        QMessageBox.information(parent, "成功", "客戶部位表寄信完成！")
    except Exception as e:
        output_text_edit.append(f"✗ 執行寄信程式時發生錯誤：{e}\n")
        QMessageBox.critical(parent, "錯誤", f"執行寄信程式時發生錯誤：{e}")

def clear_output_window(output_text_edit, parent):
    """清除輸出視窗"""
    output_text_edit.clear()
    output_text_edit.append("輸出視窗已清除\n")
# -*- coding: utf-8 -*-
"""
Created on Thu May 23 16:36:08 2024

@author: WAYNE.HUANG
"""

import pandas as pd
import win32com.client as win32
import datetime
import os
import time

os.system('taskkill /im outlook.exe /f')

tday = datetime.datetime.today().strftime('%Y%m%d')

df = pd.read_excel(r'\\10.72.228.112\cbas業務公用區\!!!交易作業區!!!\議價交易\議價交易3.0.xlsm', sheet_name='寄信', header=[1])
df = df.iloc[:, 3:]

df_bargaining = df.iloc[:, :11]

df_bargaining_buy = df_bargaining[df_bargaining['統一證買進/賣出'] == '買進']  #統一證買進
if len(df_bargaining_buy) > 0:
    df_t0_buy = df_bargaining_buy[df_bargaining_buy['交割日'].str.contains('T\+0', na=False)] #T+0買進
    df_t1_buy = df_bargaining_buy[df_bargaining_buy['交割日'].str.contains('T\+1', na=False)] #T+1買進
    qty_t0_buy = len(df_t0_buy)
    qty_t1_buy = len(df_t1_buy)
else:
    qty_t0_buy = 0
    qty_t1_buy = 0

df_bargaining_sell = df_bargaining[df_bargaining['統一證買進/賣出'] == '賣出'] #統一證賣出
if len(df_bargaining_sell) > 0:
    df_t0_sell = df_bargaining_sell[df_bargaining_sell['交割日'].str.contains('T\+0', na=False)] #T+0買進
    df_t1_sell = df_bargaining_sell[df_bargaining_sell['交割日'].str.contains('T\+1', na=False)] #T+1買進
    qty_t0_sell = len(df_t0_sell)
    qty_t1_sell = len(df_t1_sell)
else:
    qty_t0_sell = 0
    qty_t1_sell = 0

df_executed = df.iloc[:, 17:] #實物履約
df_executed_t1 = df_executed[df_executed['交割'] == 'T+1'] #T+1
df_executed_t2 = df_executed[df_executed['交割'] == 'T+2'] #T+2

df_bank = df.iloc[:, 14:17] #銀行資料
df_bank = df_bank[df_bank['客戶'].notna()]

qty_executed_t1 = len(df_executed_t1)
qty_executed_t2 = len(df_executed_t2)

t0_total = qty_t0_buy + qty_t0_sell
t1_total = qty_t1_buy + qty_t1_sell
t0_t1_total = t0_total + t1_total
executed_total = qty_executed_t1 + qty_executed_t2


#======================Mail======================

df_t0_buy_html = df_t0_buy.to_html(index=False) if 'df_t0_buy' in globals() else ''
df_t0_sell_html = df_t0_sell.to_html(index=False) if 'df_t0_sell' in globals() else ''
df_t1_buy_html = df_t1_buy.to_html(index=False) if 'df_t1_buy' in globals() else ''
df_t1_sell_html = df_t1_sell.to_html(index=False) if 'df_t1_sell' in globals() else ''

outlook = win32.Dispatch("Outlook.Application")

mail = outlook.CreateItem(0)

sendacc = None
for account in outlook.Session.Accounts:
    if account.DisplayName == 'PSC.CBAS@uni-psg.com':
        sendacc = account
        break

mail._oleobj_.Invoke(*(64209,0,8,0,sendacc))
mail.Subject = f'{tday}__{t0_total}筆議價交易T+0，{t1_total}筆議價交易T+1，{executed_total}筆實物履約T+1_附件'
email = 'vanassa@uni-psg.com;12267@uni-psg.com;MIKE@uni-psg.com;DANIEL02@uni-psg.com;YUNA.WU@uni-psg.com;IRENELIN@uni-psg.com;10176@uni-psg.com;CHUN-HUEI@uni-psg.com;AMMYCHANG@uni-psg.com;MEILAN@uni-psg.com;ERICCHEN@uni-psg.com;CATHERINE@uni-psg.com;GRACE.ROSA@uni-psg.com;XX24923051@uni-psg.com;NBDCHANG@uni-psg.com;IRENEHUANG@uni-psg.com;VANASSA@uni-psg.com;CHARLESP@uni-psg.com;PANGYEN@uni-psg.com;YIHUI@uni-psg.com;YUCHIN.HSUEH@uni-psg.com;LINDY00@uni-psg.com;P5480@uni-psg.com;EMMA@uni-psg.com;95105@uni-psg.com;CHANTAL.CHU@uni-psg.com;YITAN9593@uni-psg.com;YICIH@uni-psg.com;KMJUI.TSAI@uni-psg.com'
content = f'''<p>Dear All,</p>

<p><u><b>議價交易T+0交割，<span style="color:Red">{t0_total}筆</span><b></u></p>

{df_t0_buy_html}
{df_t0_sell_html}

<p><u><b>議價交易T+1交割，<span style="color:Red">{t1_total}</span>筆<b></u></p>

{df_t1_buy_html}
{df_t1_sell_html}

<p><u><b>實物履約T+1交割，<span style="color:Red">{executed_total}</span>筆<b></u></p>

{df_executed_t1.to_html(index=False)}

<p><u><b>交割資訊<b></u></p>

{df_bank.to_html(index=False)}
'''

mail.To = email
mail.HTMLBody = content

attachment = rf'\\10.72.228.112\cbas業務公用區\!!!交易作業區!!!\議價交易\議價交易內部通知\議價交易_{tday}.pdf'
try:
    mail.Attachments.Add(attachment)
except:
    pass

#Mail Setting  

#mail.Display(True)
mail.Send()

time.sleep(5)

#os.system('taskkill /im outlook.exe /f')

print('議價交易通知已寄出')
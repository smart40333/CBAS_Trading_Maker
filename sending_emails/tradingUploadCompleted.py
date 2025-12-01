# -*- coding: utf-8 -*-
"""
Created on Wed May 22 14:18:21 2024

@author: WAYNE.HUANG
"""

import pandas as pd
import datetime
import win32com.client as win32
import time
import os

os.system('taskkill /im outlook.exe /f')

df = pd.read_excel(r'\\10.72.228.112\cbas業務公用區\!!!交易作業區!!!\客戶部位表(對帳單)\庫存部位\7.統一證CBAS各式交易確認書.xlsm', sheet_name='客戶當日成交', usecols='A', header=None)
content = df.iloc[0,0]

tday = datetime.datetime.today().strftime('%Y%m%d')

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

sendacc = None

for account in outlook.Session.Accounts:
    if account.DisplayName == 'PSC.CBAS@uni-psg.com':
        sendacc = account
        break

mail._oleobj_.Invoke(*(64209,0,8,0,sendacc))
mail.Subject = f'CBAS本日交易_{tday}' 

email_content = f'Dear All,\n{content}(請留意可能有扣款)；\n\n以上，再麻煩協助後續作業，謝謝。'
email_list = 'MIKE@uni-psg.com;IRENELIN@uni-psg.com;10176@uni-psg.com;AMMYCHANG@uni-psg.com;MEILAN@uni-psg.com;CATHERINE@uni-psg.com;GRACE.ROSA@uni-psg.com;NBDCHANG@uni-psg.com;WAYNE.HUANG@uni-psg.com;CHARLESP@uni-psg.com;PANGYEN@uni-psg.com;LINDY00@uni-psg.com;EMMA@uni-psg.com;yuna.wu@uni-psg.com;chun-huei@uni-psg.com;P5480@uni-psg.com;95105@uni-psg.com;CHANTAL.CHU@uni-psg.com;YITAN9593@uni-psg.com;12267@uni-psg.com;vanassa@uni-psg.com;KMJUI.TSAI@uni-psg.com'

mail.Body = email_content
mail.HTMLBody
#mail.HTMLBody = "<h2>HTML Message body</h2>"

#Mail Setting  
mail.To = email_list

#mail.Display(True)
mail.Send()
time.sleep(5)

#os.system('taskkill /im outlook.exe /f')
print('本日交易已寄出')


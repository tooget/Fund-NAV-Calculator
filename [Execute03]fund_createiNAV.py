#-*- coding:utf-8 -*-

from Modules import Fund_settings
from Modules import Fund_modules
import time

# 쿼리 조회기준일자 입력하세요
#datelist = Fund_modules.general_func.create_datelist(2,'query')
dateToCreate = str(input("Input Date to Create NAV records : "))
datelist = [dateToCreate]

# 쿼리 조회 기준 펀드코드
fundcodes = Fund_settings.codeAvailable(datelist[0])


print("FundNavSummary, FundDeriPnl records creation")

# 실행 --------------------------------------------------------------------------------------------------------------------
fundQuery = Fund_modules.fund.fund_createiNAV()
i = 0
for date in datelist:
    for fundcode in fundcodes:
        i = i + 1
        print(i,'/',len(fundcodes))
        fundQuery.FundNavSummary(fundcode['<declassified>'], date)
        fundQuery.FundDeriPnl(fundcode['<declassified>'], date)
# ------------------------------------------------------------------------------------------------------------------------
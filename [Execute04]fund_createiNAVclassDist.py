#-*- coding:utf-8 -*-

from Modules import Fund_settings
from Modules import Fund_modules
import time

# 처리 메소드 클래스 선언 ------------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------------------------------------

# 쿼리 조회기준일자 입력하세요
#datelist = Fund_modules.general_func.create_datelist(2,'query')
dateToCreate = str(input("Input Date to Create NAV records : "))
datelist = [dateToCreate]

# 쿼리 조회 기준 펀드코드
fundCode_motherChildPairs = Fund_settings.codePairs_mothernchild(datelist[0])
fundCode_childrenPairs = Fund_settings.codePairs_children(datelist[0])




print("FundNavSummary Class Funds PnL Distribution")
print('-'*50)
startTimeStamp = time.time()
print("START : ",time.ctime(startTimeStamp))
print("-"*50)


# 실행 1/3 -------------------------------------------------------------------------------------------------------------------
fundQuery0 = Fund_modules.fund.fund_createiNAV()
for date in datelist:
    fundQuery0.FundNavSummary_classtypedist0(date)
#------------------------------------------------------------------------------------------------------------------------

print('-'*50)

# 실행 2/3 -------------------------------------------------------------------------------------------------------------------
fundQuery1 = Fund_modules.fund.fund_createiNAV()
i = 0
for date in datelist:
    for motherCode in fundCode_childrenPairs.keys():
        i = i + 1
        print(i,'/',len(fundCode_childrenPairs.keys()))
        #print(fundcode['mother'],fundcode['child'],date)
        fundQuery1.FundNavSummary_classtypedist1(motherCode,fundCode_childrenPairs[motherCode],date)
# ------------------------------------------------------------------------------------------------------------------------

print('-'*50)

# 실행 3/3 -------------------------------------------------------------------------------------------------------------------
fundQuery2 = Fund_modules.fund.fund_createiNAV()
j = 0
for date in datelist:
    for fundcode in fundCode_motherChildPairs:
        j = j + 1
        print(j,'/',len(fundCode_motherChildPairs))
        #print(fundcode['mother'],fundcode['child'],date)
        fundQuery2.FundNavSummary_classtypedist2(fundcode['<declassified>'],fundcode['<declassified>'],date)
# ------------------------------------------------------------------------------------------------------------------------

print('-'*50)

endTimeStamp = time.time()
print("-"*50)
print("END : ",time.ctime(endTimeStamp))
print("Running Time : ",round(endTimeStamp-startTimeStamp,2))
print("-"*50)
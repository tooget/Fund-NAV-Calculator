#-*- coding:utf-8 -*-

from Modules import Fund_deri_modules

b = Fund_deri_modules.bls_cal()
d = Fund_deri_modules.data_struc()

P_date = str(input("Pricing Reference Date : "))


print('\n','Ref Option Price Implied Volatility Comparison','-'*50)
verifyer = 0
k = 0
items = d.refopprice_targetTocompare(P_date)

est_price = lambda x: 1e-10 if x == 0 else x

print(
    '펀드코드', '|',
    '종목', '|',
    '블룸버그 종목명', '|',
    'Moneyness', '|',
    '외부평가가격(REF)', '|',
    '내부평가가격', '|',
    '내부평가가격(검증)', '|',
    '내부평가 변동성', '|',
    '외부평가 변동성',
    )

for item in items:

    price = b.bls_op_price(
            item['구분'][0],
            float(item['기초자산_종가'])-float(item['예상배당']),
            float(item['행사가']),
            d.calDaysRemained(P_date,item['만기일']),
            float(item['무위험이자율']),
            float(item['변동성']),
            0,
            )

    refiv = b.bls_ivol(
                item['구분'][0],
                float(item['기초자산_종가'])-float(item['예상배당']),
                float(item['행사가']),
                d.calDaysRemained(P_date,item['만기일']),
                float(item['무위험이자율']),
                0,
                est_price(float(item['평가가격'])),
                )
    print(
        item['<declassified>'], '|',
        item['종목'].strip(), '|',
        item['블룸버그_종목명'].strip(), '|',
        round(d.calMoneyness(float(item['행사가']),float(item['기초자산_종가']))/(d.calDaysRemained(P_date,item['만기일'])**(0.5)),4), '|',
        round(float(item['평가가격']),4), '|',
        round(float(item['자체평가가격']),4), '|',
        round(price,4), '|',
        round(float(item['변동성']),4), '|',
        round(float(refiv),4),
        )
    verifyer = verifyer + round(float(item['자체평가가격'])-price,9)
if verifyer == 0.0 : print('OK', '-'*50)
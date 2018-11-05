#-*- coding:utf-8 -*-

class bls_cal:

    def bls_op_price(self,flag,S,K,T,r,v,q):
        # The generalized Black and Scholes formula
        import math
        from scipy.stats import norm

        n = norm.pdf
        N = norm.cdf
        b = r-q

        d1 = (math.log(S/K)+(b+v*v/2.)*T)/(v*math.sqrt(T))
        d2 = d1-v*math.sqrt(T)
        #print(d1)
        #print(N(d1))    

        if flag.lower() in ('c'):
            op_cprice = S*math.exp((b-r)*T)*N(d1)-K*math.exp(-r*T)*N(d2)
            return(op_cprice)
        elif flag.lower() in ('p'):
            op_pprice = K*math.exp(-r*T)*N(-d2)-S*math.exp((b-r)*T)*N(-d1)
            return(op_pprice)
        elif flag.lower() in ('vega'):
            op_vega = S*math.exp((b-r)*T)*N(d1)*math.sqrt(T)
            return(op_vega)

    def bls_ivol(self,flag,S,K,T,r,q,est_price):
        ivol = 0.5
        i = 0
        while True:
            f = self.bls_op_price(
                    flag,
                    S,
                    K,
                    T,
                    r,
                    ivol,
                    q,
                    ) - est_price
            if i > 15 and f < 0.00001:
                break
            fprime = self.bls_op_price(
                    'vega',
                    S,
                    K,
                    T,
                    r,
                    ivol,
                    q,
                    )
            ivol = ivol - f / fprime
            i = i + 1
        #print(i)
        return(ivol)

        
class data_struc:

    # MySQL DB connection -----------------------------------------------------------------------------------------------------
    import pymysql
    conv = pymysql.converters.conversions.copy()    #pymysql.converters.conversions : fetch return에 대한 datatype 기본 세팅 복사
    conv[10] = str                                  #pymysql query return시 table 상 date type field 값을 문자로 리턴함. 기본값은 datetime.date(yyyy,mm,dd)
    con = pymysql.connect(
            host='<declassified>',
            port='<declassified>',
            user='<declassified>',
            passwd='<declassified>',
            db='<declassified>',
            charset='utf8',
            conv = conv,                            #pymysql 연결시 return 세팅을 변경함.
            )    
    cs = con.cursor(pymysql.cursors.DictCursor)
    print('\n','%s is connected' % 'MySQL fund db','\n')
    # ------------------------------------------------------------------------------------------------------------------------

    def calDaysRemained(self,P_date,K_date):
        import datetime as dt
        daysleft = dt.datetime.strptime(K_date,'%Y-%m-%d') - dt.datetime.strptime(P_date,'%Y-%m-%d')
        T = daysleft.days
        #print(T)
        T = float(T) / 365.
        return(T)

    def calMoneyness(self,K,S):
        import math
        moneyness = math.log(K/S)
        return(moneyness)

    def refopprice_targetTocompare(self, P_date):
        import json
        setDate = "SET @<declassified> = '" + P_date + "';"
        mainQuery = self.ReadSqlFile('<declassified>')
        self.cs.execute(setDate)
        self.cs.execute(mainQuery)
        temp = self.cs.fetchall()
        return(temp)

    def ReadSqlFile(self, filename):
        path = '<declassified>'
        filepath = path + filename + '.sql'
        query = open(filepath, 'r').read()
        return(query)
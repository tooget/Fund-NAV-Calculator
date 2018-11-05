#-*- coding:utf-8 -*-

from Modules import Fund_settings

class general_func:
    #요약 : 여러 파일에 중복사용된 함수를 모아놓은 클래스
    def create_datelist(before_Ndays_from_today,args):
        #요약 : 날짜의 목록을 생성하는 함수
        #input : 숫자(오늘로부터 며칠 전까지의 리스트를 만들 것인지)
        #input(옵션) : filename or Else
        #return : 과거부터 어제까지의 날짜(str/YYYY-MM-DD 또는 YYYYMMDD)의 리스트
        import datetime as dt
        temp = dt.datetime.today()
        strftime_type = '%Y%m%d' if args == 'filename' else '%Y-%m-%d'
        datelist = [(temp-dt.timedelta(days=x)).strftime(strftime_type) for x in range(0,before_Ndays_from_today)]
        datelist.sort(reverse=True)
        datelist = datelist[1:]
        return(datelist)
    def filesInDirectory(directory_path):
        #요약 : 폴더 내의 정렬된 파일 목록을 가져오는 함수
        #input : 확장자 포함 폴더 경로
        #return : 특정 확장자를 가진 파일 경로들의 리스트
        import glob
        filesList = glob.glob(directory_path)
        filesList.sort()
        return(filesList)

class fund:
    class fund_extractXLSXfromFAS:
        #요약 : gui 노가다 자동화를 위한 함수들
        import pyautogui as pyg
        import pyperclip as cl
        import time
        def checkImagesOnscreenLoop(self,filename,args,loopCount):
            #요약 : 조건을 충족할 때까지 화면에서 특정 이미지를 반복하여 탐색함.
            #input : 경로 포함된 파일명, 탐색옵션('exist' or 'notexist'), 반복 탐색 횟수
            #return : 없음.
            #참조1 : 이미지를 탐색하는 함수(checkImageOnscreen)을 참조함.
            #참조2 : 반복 탐색 횟수 초과시 TypeError
            i = 0
            while self.checkImageOnscreen(filename,args) is False:
                i = i + 1
                #print(i)
                if i > loopCount:
                    raise(TypeError('STOP : Counts over.'))
            self.time.sleep(0.5)
        def checkImageOnscreen(self,filename,args):
            #요약 : 화면에서 특정 이미지가 있는지(혹은 없는지) 확인함.
            #input : 경로 포함된 파일명, 탐색옵션('exist' or 'notexist')
            #return : True or False
            #참조 : 이미지를 탐색을 반복하는 함수(checkImagesOnscreenLoop)에 참조됨.
            if args == 'exist':
                if type(self.pyg.locateOnScreen(filename)) == tuple:
                    #print('Pass : Found the %s image on this screen.' % filename.split('\\')[-1])
                    return(True)
                elif type(self.pyg.locateOnScreen(filename)) == type(None):
                    print('Fail : There is no %s image on this screen.' % filename.split('\\')[-1])
                    return(False)
            elif args == 'notexist':
                if type(self.pyg.locateOnScreen(filename)) == tuple:
                    print('Fail : Found the %s image on this screen.' % filename.split('\\')[-1])
                    return(False)
                elif type(self.pyg.locateOnScreen(filename)) == type(None):
                    #print('Pass : There is no %s image on this screen.' % filename.split('\\')[-1])
                    return(True)
        def imageCenterClick(self,filename):
            #요약 : 화면에서 찾은 특정이미지의 가운데를 클릭함.
            #input : 경로 포함된 파일명
            #return : 없음.
            temp_set = self.pyg.locateOnScreen(filename)
            self.pyg.click(x=temp_set[0]+temp_set[2]/2,y=temp_set[1]+temp_set[3]/2)
        def imageRightSideClick(self,filename):
            #요약 : 화면에서 찾은 특정이미지에서 약간 떨어진 오른쪽 부분을 클릭함.
            #input : 경로 포함된 파일명
            #return : 없음.
            temp_set = self.pyg.locateOnScreen(filename)
            self.pyg.click(x=temp_set[0]+temp_set[2]+temp_set[2]/3,y=temp_set[1]+temp_set[3]/2)
        def imageCenterDoubleClick(self,filename):
            #요약 : 화면에서 찾은 특정이미지의 가운데를 더블클릭함.
            #input : 경로 포함된 파일명
            #return : 없음.
            temp_set = self.pyg.locateOnScreen(filename)
            self.pyg.doubleClick(x=temp_set[0]+temp_set[2]/2,y=temp_set[1]+temp_set[3]/2)
        def inputKeys(self,string):
            #요약 : 입력받은 문자열을 키 입력함.
            #input : 문자열(str)
            #return : 없음.
            self.pyg.press(string)
        def inputHotkeys(self,key1,key2):
            #요약 : 단축키 조합을 키 입력함.
            #input : 단축키1(str), 단축키2(str)
            #return : 없음.
            self.pyg.hotkey(key1,key2)
        def pasteKey(self,string):
            #요약 : 입력받은 문자열을 붙여넣기함.
            #input : 문자열(str)
            #return : 없음.
            self.cl.copy(string)
            self.pyg.hotkey('ctrl','v')
    class fund_createiNAV:
        # MySQL DB connection -----------------------------------------------------------------------------------------------------
        from sqlalchemy import create_engine
        import pymysql
        pymysql.install_as_MySQLdb()
        import MySQLdb
        #DB Charset이 utf8, Collation Name이 utf8_general_ci 인지 확인할 것!
        #pandas Dataframe.to_sql() 메소드가 pymysql의 connection으로는 연결되지 않음.
        engine = create_engine('mysql+mysqldb://'
                        + Fund_settings.fundDB_user +':'
                        + Fund_settings.fundDB_passwd +'@'
                        + Fund_settings.fundDB_host + '/'
                        + Fund_settings.fundDB_name + '?charset=utf8', encoding='utf8')
        conn = engine.connect()
        print('\n','%s is connected' % 'MySQL fund db','\n')
        # ------------------------------------------------------------------------------------------------------------------------
        con = pymysql.connect(
                host=Fund_settings.fundDB_host,
                port=Fund_settings.fundDB_port,
                user=Fund_settings.fundDB_user,
                passwd=Fund_settings.fundDB_passwd,
                db=Fund_settings.fundDB_name,
                charset=Fund_settings.fundDB_charset,
                )
        cs = con.cursor()
        # ------------------------------------------------------------------------------------------------------------------------
        def FundNavSummary(self, fundCode_target, dateStr) :
            #요약 : <declassified>에 펀드별, 기준일자별 레코드 생성
            #input: 펀드코드, 기준일, sql파일(Fund_settings.iNAVqueryFilename)
            #return: 없음.
            #참고1 : 여러 다른 테이블의 데이터를 합쳐서 조회함. sql파일 참고
            #참고2 : Dataframe으로 쿼리 결과를 가공하여 신규 레코드 입력하도록 동작함.

            import pandas as pd
            import datetime
            #기준 펀드코드 변수 선언 쿼리 세팅
            setFundcodeQuery = "SET @<declassified> = '" + fundCode_target + "';"
            #기준일자 변수 선언 쿼리 세팅
            setDateQuery = "SET @<declassified> = '" + dateStr + "';"
            #내부 기준가 메인쿼리 SQL파일 불러오기/세팅
            mainQuery = self.ReadSqlFile(Fund_settings.iNAVqueryFilename)
            #기준 펀드코드 쿼리 실행
            self.cs.execute(setFundcodeQuery)
            #기준일자 쿼리 실행
            self.cs.execute(setDateQuery)
            #read_sql로 내부기준가 메인쿼리를 실행하여 dataframe을 리턴받음.
            df = pd.read_sql(mainQuery, self.con)
            #<declassified>, Date 칼럼 삭제
            df = df.drop(['<declassified>', '<declassified>'], axis = 1)
            #펀드코드, 날짜, 연산시간을 레코드로 삽입
            cols = list(df.columns.values)
            vals = [['<declassified>',fundCode_target], ['<declassified>',dateStr], ['<declassified>', str(datetime.datetime.now())]]
            df = df.append([dict(zip(cols,val)) for val in vals], ignore_index=False)
            #빈 행 제거 및 인덱스 초기화
            df = df.fillna(0)
            df = df.reset_index(drop=True)
            #행렬 전환
            df = df.transpose()
            #첫번째 행을 dataframe 헤더로 바꿈
            new_header=df.iloc[0]
            df = df.rename(columns=new_header)
            df = df[1:]
            try:
                #<declassified> 테이블에 레코드 추가
                df.to_sql(con=self.conn, name='<declassified>', if_exists='append', index_label='<declassified>', index=False)
                print('%s 의 %s 일자 <declassified> 레코드가 입력되었습니다.' % (fundCode_target, dateStr))
            except self.pymysql.err.InternalError:
                #테이블 스키마가 맞지 않으면 InternalError가 뜸.
                print('<declassified> 맞지 않는 필드가 있습니다.' % (fundCode_target, dateStr))
            except:
                print('%s 의 %s 일자 <declassified> 레코드 입력시 에러가 발생하였습니다..' % (fundCode_target, dateStr))
        def FundDeriPnl(self, fundCode_target, dateStr) :
            #요약 : <declassified>에 펀드코드, 기준일자, 파생종목별 내부가격/손익비교 레코드를 생성함.
            #input : 펀드코드, 기준일자, sql파일(Fund_settings.DeriPnLqueryFilename)
            #return : 없음.
            #참고1 : 여러 다른 테이블의 데이터를 합쳐서 조회함. sql파일 참고
            #참고2 : Dataframe으로 쿼리 결과를 가공하여 신규 레코드 입력하도록 동작함.

            import pandas as pd
            import datetime
            #펀드코드 변수 선언 쿼리 세팅
            setFundcodeQuery = "SET @<declassified> = '" + fundCode_target + "';"
            #기준일자 변수 선언 쿼리 세팅
            setDateQuery = "SET @<declassified> = '" + dateStr + "';"
            #내용조회쿼리 파일 불러오기
            mainQuery = self.ReadSqlFile(Fund_settings.DeriPnLqueryFilename)
            #펀드코드 변수 선언 쿼리 실행
            self.cs.execute(setFundcodeQuery)
            #기준일자 변수 선언 쿼리 실행
            self.cs.execute(setDateQuery)
            #내용조회쿼리 dataframe으로 실행
            df = pd.read_sql(mainQuery, self.con)
            #dataframe에 펀드코드 추가
            df['<declassified>'] = fundCode_target
            #dataframe에 연산시간 추가
            df['<declassified>'] = str(datetime.datetime.now())
            #빈 행 삭제
            df = df.fillna(0)
            try:
                df.to_sql(con=self.conn, name='<declassified>', if_exists='append', index_label='<declassified>', index=False)
                print('%s 의 %s 일자 <declassified> 레코드가 입력되었습니다.' % (fundCode_target, dateStr))
            except:
                print('%s 의 %s 일자 <declassified> 레코드 입력시 에러가 발생하였습니다..' % (fundCode_target, dateStr))
        def FundNavSummary_classtypedist0(self, dateStr) :
            #요약 : <declassified>의 각 클래스의 조정좌수 필드를 재계산하여 업데이트 쿼리를 실행함.
            #input: 기준일
            #return: 없음

            setDate = "SET @<declassified> := '" + dateStr + "';"
            #print('setDate : ',setDate)
            setTargetFund = "Set @<declassified> := (select concat('<declassified>', <declassified>) from <declassified> where <declassified> is not NULL and <declassified> like '<declassified>' and <declassified> > <declassified>);"
            #print('setTargetFund : ',setTargetFund)
            setValueQuery = "SET @<declassified> := (select <declassified> from <declassified> where <declassified> = @<declassified> and <declassified> = @<declassified>);"
            #print(setValueQuery)
            mainQuery = "Update <declassified> set <declassified> = @<declassified> where <declassified> = @<declassified> and <declassified> = @<declassified>;"
            #print(mainQuery)
            self.cs.execute(setDate)
            #print(setDate,' is success')
            self.cs.execute(setTargetFund)
            #print(setTargetFund,' is success')
            self.cs.execute(setValueQuery)
            #print(setValueQuery,' is success')
            self.cs.execute(mainQuery)
            #print(mainQuery,' is success')
            self.con.commit()
            print('<declassified>의 클래스펀드 조정좌수 업데이트가 완료되었습니다.')
        def FundNavSummary_classtypedist1(self, motherFundCode, childrenFundCodeList, dateStr) :
            #요약 : <declassified>의 운용펀드의 조정좌수 필드를 재계산하여 업데이트 쿼리를 실행함.
            #input: 기준일, 운용펀드코드, 클래스펀드코드
            #return: 없음
            #참고 : FundNavSummary_classtypedist0을 실행한 뒤 FundNavSummary_classtypedist1를 실행해야 함.

            setDate = "SET @<declassified> := '" + dateStr + "';"
            #print('setDate : ',setDate)
            childrenStr = "','".join(str(x) for x in childrenFundCodeList)
            #print('childrenStr : ',childrenStr)
            setValueQuery = "SET @<declassified> := (select sum(<declassified>) from <declassified> where <declassified> = '" + dateStr + "' and <declassified> in ('" + childrenStr + "'));"
            #print('setValueQuery : ',setValueQuery)
            setMotherFundcode = "SET @<declassified> := '" + motherFundCode + "';"
            #print('setMotherFundcode : ',setMotherFundcode)
            mainQuery = 'update <declassified> set <declassified> = @<declassified> where <declassified> = @<declassified> and <declassified> = @<declassified>;'
            #print('mainQuery : ',mainQuery)
            self.cs.execute(setDate)
            #print(setDate,' is success')
            self.cs.execute(setValueQuery)
            #print(setValueQuery,' is success')
            self.cs.execute(setMotherFundcode)
            #print(setMotherFundcode,' is success')
            self.cs.execute(mainQuery)
            #print(mainQuery,' is success')
            self.con.commit()
            print('<declassified>에서 ',motherFundCode,'의 조정좌수 업데이트가 완료되었습니다.')
        def FundNavSummary_classtypedist2(self, motherFundCode, childrenFundCodeList, dateStr) :
            #요약 : <declassified>상 운용/클래스펀드의 조정좌수를 기준으로, 하위클래스에 손익 필드를 분배하여 새로운 레코드를 생성하고, 기존의 레코드를 삭제함.
            #input: 기준일, 운용펀드코드, 클래스펀드코드
            #return: 없음
            #참고1 : FAS상 하위클래스의 운용손익은, 운용클래스 지분 소유 개념으로 반영되어 있어 구분이 되지 않음. FAS8141 참고
            #참고2 : FundNavSummary_classtypedist1을 실행한 뒤 FundNavSummary_classtypedist2를 실행해야 함.
            #참고3 : 여러 다른 테이블의 데이터를 합쳐서 조회함. sql파일 참고
            #참고4 : Dataframe으로 쿼리 결과를 가공하여 신규 레코드 입력하도록 동작함.

            import pandas as pd
            import datetime
            setDate = "SET @<declassified> := '" + dateStr + "';"
            #print(setDate)
            setMotherFundcode = "SET @<declassified> := '" + motherFundCode + "';"
            #print(setMotherFundcode)
            setChildFundcode = "SET @<declassified> := '" + childrenFundCodeList + "';"
            #print(setChildFundcode)
            deleteChildRecord = "Delete from <declassified> where <declassified> = '" + dateStr + "' and <declassified> = '" + childrenFundCodeList +"' and (abs(<declassified>)+abs(<declassified>)+abs(<declassified>)) = <declassified>;"
            #print(deleteChildRecord)
            mainQuery = self.ReadSqlFile(Fund_settings.iNAVclassfundQueryFilename)
            #print(mainQuery)
            self.cs.execute(setDate)
            #print('setDate')
            self.cs.execute(setMotherFundcode)
            #print('setMotherFundcode')
            self.cs.execute(setChildFundcode)
            #print('setChildFundcode')
            df = pd.read_sql(mainQuery, self.con)
            #print('read Query result')
            df['<declassified>'] = str(datetime.datetime.now())
            df = df.fillna(0)
            #print(df)
            #print('pd setting complete')
            self.cs.execute(deleteChildRecord)
            self.con.commit()
            #print('deleted')
            try:
                df.to_sql(con=self.conn, name='<declassified>', if_exists='append', index_label='<declassified>', index=False)
                print('%s 의 %s 일자 <declassified> 클래스 배분내용 레코드가 삭제 후 신규 입력되었습니다.' % (childrenFundCodeList, dateStr))
            except:
                print('%s 의 %s 일자 <declassified> 클래스 배분내용 레코드 삭제 후 신규 입력시 에러가 발생하였습니다..' % (childrenFundCodeList, dateStr))
        def ReadSqlFile(self, filename):
            #요약 : sql 파일을 str으로 읽어줌.
            #input : sql 파일명, sql 파일 경로(Fund_settings.fundQueryPath)
            #output : 쿼리(str)
            filepath = Fund_settings.fundQueryPath + filename + '.sql'
            query = open(filepath, 'r').read()
            return(query)
    class fund_updateXLSXtoMySQLDB:
        #fund_updateXLSXtoMySQLDB 모듈 개요
        #input: 파일경로, Fund_settings.codeofMotherFunds_SpecificHeaderOnly()
        #return : DB에 넣기 위한 Dataframe

        fundcode_list = None
        def futopprice(self,filename_with_path,filename_without_path):
            #요약 : <declassified>의 내용을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명, Fund_settings.codeofMotherFunds_SpecificHeaderOnly()
            #return: DB에 넣을 운용파일 내용(dataframe)
            #참고 : <declassified>의 <declassified>를 가져오기 위해 DB를 조회함.(Fund_settings.codeofMotherFunds_SpecificHeaderOnly())

            # 운용펀드코드 조회
            self.fundcode_list = list( zip(*(y.values() for y in Fund_settings.codeofMotherFunds_SpecificHeaderOnly())) )[0]
            import pandas as pd
            from datetime import datetime
            # 운용파일 읽어오기
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #'<declassified>' 파일에서 필요한 내용을 추출하기 위한 Dataframe 가공 절차
            #Step1. <declassified>열에서 '<declassified>' 값의 위치와 '<declassified>' 값의 위치를 찾는다.
            row_count_start = data[data[0] == 'BM'].index.tolist()
            row_count_end = data[data[0] == '합계(BM대비)'].index.tolist()
            #Step2. <declassified> 값의 위치를 순서쌍 'range_list'로 묶는다. 이 사이의 행을 추출하면 됨.
            range_list = [(row_count_start[x],row_count_end[x]) for x in range(0,len(row_count_start))]
            #Step3. 빈 데이터프레임을 만들고, 그 안에 순서쌍 구간의 내용을 정리하여 추가함.
            i = 0
            data_refined = pd.DataFrame()
            for range_idx in range_list:
                #행 자르기
                temp = data[range_idx[0]-1:range_idx[1]]
                #<declassified>열 이후의 열을 삭제
                temp = temp.iloc[:,1:37]
                #<declassified>, <declassified>열 삭제
                temp = temp.drop(temp.columns[list(range(19,27))+[28]], axis=1)
                #<declassified>열 중 비어있는 행 삭제 
                temp = temp.dropna(subset=[2],axis=0)
                #<declassified>열에 순서대로 <declassified>에서 가져온 모펀드코드추가
                temp.loc[:,0] = self.fundcode_list[i]
                #인덱스 순서 칼럼 추가
                temp.reset_index(inplace=True)
                temp.reset_index(inplace=True)
                temp = temp.drop(labels='index',axis=1)
                #빈 데이터프레임'data_refined'에 추가
                data_refined = data_refined.append(temp)
                del(temp)
                i = i + 1
            #Step5. 헤더 세팅 및 변경
            new_headers = [
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                            '<declassified>',
                        ]
            data_refined.columns = new_headers
            #Step6. 기준일 및 연산시간 입력    
            data_refined['<declassified>'] = datetime.strptime(filename_without_path.split('_')[-1][0:8],'%Y%m%d').strftime('%Y-%m-%d')
            data_refined['<declassified>'] = str(datetime.now())
            return(data_refined)
        def specificfundissueinfo(self, filename_with_path, filename_without_path):
            #요약 : <declassified> 내 <declassified> 내용을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명, Fund_settings.codeofMotherFunds_SpecificHeaderOnly()
            #return: DB에 넣을 <declassified> 정보(dataframe)
            #참고 : <declassified>의 <declassified>를 가져오기 위해 DB를 조회함.(Fund_settings.codeofMotherFunds_SpecificHeaderOnly())

            self.fundcode_list = list( zip(*(y.values() for y in Fund_settings.codeofMotherFunds_SpecificHeaderOnly())) )[0]
            import pandas as pd
            from datetime import datetime
            #파일을 pandas의 read_excel을 통해 dataframe으로 불러옴.
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #첫 레코드를 헤더로 변환
            new_header=data.iloc[0]
            data = data.rename(columns=new_header)
            data = data[1:]
            #펀드 구분 추출
            fundRefs = data['<declassified>'].dropna()
            #구분 열의 펀드구분 공백을 채워넣기
            data['<declassified>'] = data['<declassified>'].fillna(method='ffill')
            #펀드별 레코드 순서 넣기
            data_refined = pd.DataFrame()
            i = 0
            for fundRef in fundRefs:
                temp = data[(data['<declassified>'] == fundRef)]
                temp = temp.reset_index(drop=True)
                #<declassified>에 순서대로 <declassified>에서 가져온 <declassified> 모펀드코드추가
                if self.fundcode_list[i][-3:] == fundRef[1:] :
                    temp['<declassified>'] = self.fundcode_list[i]
                else :
                    temp['<declassified>'] = '<declassified>'
                temp.reset_index(level=0, inplace=True)
                data_refined = data_refined.append(temp)
                #print(temp)
                del(temp)
                i = i + 1
            #Column Name이 nan인 세 부분을 각각 순차적인 이름으로 바꿔넣기
            data_refined.rename(
                                columns={
                                    '<declassified>':'<declassified>',
                                    '<declassified>':'<declassified>',
                                    '<declassified>':'<declassified>',
                                        }
                                ,inplace=True
                                )
            new_headers = list(data_refined.columns.fillna())
            for i in range(0,len(new_headers)):
                if new_headers[i] is not None:
                    pass
                elif new_headers[i] is None:
                    new_headers[i] = new_headers[i-1][:-1] + str(int(new_headers[i-1][-1])+1)
            data_refined.columns = new_headers
            #기준일입력
            data_refined['<declassified>'] = datetime.strptime(filename_without_path.split('_')[-1][0:8],'%Y%m%d').strftime('%Y-%m-%d')
            #연산시간 입력
            data_refined['<declassified>'] = str(datetime.now())
            return(data_refined)
        def codematch(self,filename_with_path,filename_without_path):
            #요약 : <declassified>상 종목(ISIN)과 운용파일상 <declassified>을 연결하기 위한 <declassified>를 DB에 넣는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 정보(dataframe)
            #참고 : 파일 내용상 <declassified>은 수작업으로 진행함.(2017년 10월 현재)

            import pandas as pd
            from datetime import datetime
            #파일을 pandas의 read_excel을 통해 dataframe으로 불러옴.
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #첫 레코드를 헤더로 변환
            new_header=data.iloc[0]
            data = data.rename(columns=new_header)
            data = data[1:]
            #연산시간 입력
            data['<declassified>'] = str(datetime.now())
            return(data)
        def refopprice(self,filename_with_path,filename_without_path):
            #요약 : REF에서 <declassified>으로 보내주는 <declassified> 내용을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 REF <declassified> 정보(dataframe)

            import pandas as pd
            from datetime import datetime
            #파일을 pandas의 read_excel을 통해 dataframe으로 불러옴.
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #첫 행를 dataframe 헤더로 변환
            new_header=data.iloc[0]
            data = data.rename(columns=new_header)
            data = data[1:]
            #기준일입력
            data['<declassified>'] = datetime.strptime(filename_without_path.split('_')[-1][0:8],'%Y%m%d').strftime('%Y-%m-%d')
            #연산시간 입력
            data['<declassified>'] = str(datetime.now())
            return(data)
        def refieprice(self,filename_with_path,filename_without_path):
            #요약 : REF에서 <declassified>으로 보내주는 <declassified> 평가내용을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 REF <declassified> 평가정보(dataframe)

            import pandas as pd
            from datetime import datetime
            #매칭시킬 <declassified> 펀드코드 리스트를 DB에서 불러옴.(REF에서 <declassified> 펀드코드를 같이 보내주지 않음.)
            self.fundcode_list = list( zip(*(y.values() for y in Fund_settings.codeofMotherFunds_SpecificHeaderOnly())) )[0]
            #파일을 pandas의 read_excel을 통해 dataframe으로 불러옴.
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #첫 행를 dataframe 헤더로 변환
            new_header=data.iloc[0]
            data = data.rename(columns=new_header)
            data = data[1:]
            #텍스트 공백 없애기
            data['<declassified>'] = data['<declassified>'].str.strip()
            #순서대로 펀드코드 추가하기
            fundcode_count_for_ie = self.fundcode_list[:len(data)]
            data['<declassified>'] = fundcode_count_for_ie
            #기준일입력
            data['<declassified>'] = datetime.strptime(filename_without_path.split('_')[-1][0:8],'%Y%m%d').strftime('%Y-%m-%d')
            #연산시간 입력
            data['<declassified>'] = str(datetime.now())
            return(data)
        def fas34(self, filename_with_path, filename_without_path):
            #요약 : <declassified>에서 엑셀다운받은 파일을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 내용(dataframe)
            #Error message: sqlalchemy.exc.DataError: (pymysql.err.DataError) (1264, "Out of range value for column '<declassified>' at row 1")
            #Trouble Shooting : 생성된 <declassified>의 'f<declassified>' 를 BIGINT(20)에서 VARCHAR(20)으로 조정후 재입력하면 됨.
            return(self.codematch(filename_with_path,filename_without_path))
        def fas41(self, filename_with_path, filename_without_path):
            #요약 : <declassified>에서 엑셀다운받은 파일을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 내용(dataframe)

            import pandas as pd
            from datetime import datetime
            #파일을 pandas의 read_excel을 통해 dataframe으로 불러옴.
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #첫 행를 dataframe 헤더로 변환
            new_header=data.iloc[0]
            data = data.rename(columns=new_header)
            data = data[1:]
            #연산시간 입력
            data['<declassified>'] = str(datetime.now())
            return(data)
        def fas800(self,filename_with_path,filename_without_path):
            #요약 : <declassified>에서 엑셀다운받은 파일을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 내용(dataframe)

            import pandas as pd
            from datetime import datetime
            #FAS 8001화면 엑셀 추출파일 불러오기
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #한줄 기준으로 정렬하기 위해 필요한 row_counts 연산
            modified_row_counts = int(pd.to_numeric(data[0],'corece').max()+2)
            #두줄 레코드를 한줄 레코드로 재정렬
            data = pd.DataFrame(data.values.reshape(modified_row_counts,-1))
            #빈 열 삭제
            data = data.dropna(axis='columns',how='all')
            #첫 레코드를 헤더로 변환하고 마지막 합계 레코드를 삭제하여 DB에 입력가능한 상태로 정리함.
            new_header=data.iloc[0] 
            data = data.rename(columns=new_header)
            data = data[1:modified_row_counts-1]
            #파일이름을 기준으로 기준일, 자체펀드코드 데이터 입력
            data['<declassified>'] = datetime.strptime(filename_without_path.split('_')[-1][0:8],'%Y%m%d').strftime('%Y-%m-%d')
            data['<declassified>'] = filename_without_path.split('_')[-2]
            #연산시간 입력
            data['<declassified>'] = str(datetime.now())
            #data = data.dropna(axis=1)
            return(data)
        def fas80(self, filename_with_path, filename_without_path):
            #요약 : <declassified>에서 엑셀다운받은 파일을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 내용(dataframe)
            return(self.refopprice(filename_with_path,filename_without_path))
        def fas81(self, filename_with_path, filename_without_path):
            #요약 : <declassified>에서 엑셀다운받은 파일을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 내용(dataframe)

            import pandas as pd
            from datetime import datetime
            #파일을 pandas의 read_excel을 통해 dataframe으로 불러옴.
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #FAS 8141화면에서 1행에 필요한 헤더 입력
            data.loc[0] = ['<declassified>','<declassified>','<declassified>','<declassified>','<declassified>','<declassified>','<declassified>','<declassified>']
            #마지막 레코드 '계정과목코드' 및 '계정과목명'에 '합계' 값 입력
            data.loc[len(data)-1] = data.loc[len(data)-1].fillna('<declassified>')
            #설정한 리스트로 'data(데이터프레임)'의 헤더 변경
            data = data.rename(columns=data.loc[0])
            #'data(데이터프레임)'의 헤더를 입력했으니 불필요한 첫번째, 두번째 행 삭제
            data = data[2:]
            #파일이름을 기준으로 기준일, 자체펀드코드 필드 추가
            data['<declassified>'] = filename_without_path.split('_')[-2]
            data['<declassified>'] = datetime.strptime(filename_without_path.split('_')[-1][0:8],'%Y%m%d').strftime('%Y-%m-%d')
            #연산시간 필드 추가
            data['<declassified>'] = str(datetime.now())
            return(data)
        def fas83(self, filename_with_path, filename_without_path):
            #요약 : <declassified>에서 엑셀다운받은 파일을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 내용(dataframe)

            import pandas as pd
            from datetime import datetime
            #파일을 pandas의 read_excel을 통해 dataframe으로 불러옴.
            data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #첫 행를 dataframe 헤더로 변환
            new_header=data.iloc[0]
            data = data.rename(columns=new_header)
            #헤더로 변환한 첫행과 합계인 마지막행 삭제
            data = data[1:-1]
            #순번 열 삭제
            data = data.drop('<declassified>', axis='columns')
            #펀드 열 펀드코드 앞에 '0' 입력 (엑셀 추출시 숫자 인식으로 앞자리 0이 지워지는 현상 보완)
            data['<declassified>'] = '<declassified>' + data['<declassified>']
            #연산시간 입력
            data['<declassified>'] = str(datetime.now())
            return(data)
        def fas88(self, filename_with_path, filename_without_path):
            #요약 : <declassified>에서 엑셀다운받은 파일을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 내용(dataframe)

            return(self.refopprice(filename_with_path,filename_without_path))
        def ef25(self, filename_with_path, filename_without_path):
            #요약 : <declassified>에서 계좌별로 받은 csv파일을 DB에 넣을 수 있는 형태로 가공하는 함수
            #input: 경로 포함 파일명, 파일명
            #return: DB에 넣을 <declassified> 내용(dataframe)

            import pandas as pd
            from datetime import datetime
            #파일을 pandas의 read_csv을 통해 dataframe으로 불러옴.
            data = pd.read_csv(filename_with_path, header=None, encoding='euc-kr')
            #data = pd.read_excel(filename_with_path, header=None, encoding='utf8', engine='xlrd')
            #첫 행를 dataframe 헤더로 변환
            new_header=data.iloc[0]
            data = data.rename(columns=new_header)
            data = data[1:]
            #파일이름을 기준으로 기준일, 자체펀드코드 필드 추가
            data['<declassified>'] = filename_without_path.split('_')[-2]
            #기준일입력
            data['<declassified>'] = datetime.strptime(filename_without_path.split('_')[-1][0:8],'%Y%m%d').strftime('%Y-%m-%d')
            #연산시간 입력
            data['<declassified>'] = str(datetime.now())
            return(data)
    class fund_updateXLSXtoMySQLDB_interface:
        # MySQL DB connection -----------------------------------------------------------------------------------------------------
        from sqlalchemy import create_engine
        import pymysql
        #from pymysql import install_as_MySQLdb
        pymysql.install_as_MySQLdb()
        import MySQLdb
        #DB Charset이 utf8, Collation Name이 utf8_general_ci 인지 확인할 것!
        #pandas Dataframe.to_sql() 메소드가 pymysql의 connection으로는 연결되지 않음.
        engine = create_engine('mysql+mysqldb://'
                        + Fund_settings.fundDB_user +':'
                        + Fund_settings.fundDB_passwd +'@'
                        + Fund_settings.fundDB_host + '/'
                        + Fund_settings.fundDB_name + '?charset=utf8', encoding='utf8')
        conn = engine.connect()
        con = pymysql.connect(
                host=Fund_settings.fundDB_host,
                port=Fund_settings.fundDB_port,
                user=Fund_settings.fundDB_user,
                passwd=Fund_settings.fundDB_passwd,
                db=Fund_settings.fundDB_name,
                charset=Fund_settings.fundDB_charset,
                )    
        cs = con.cursor()
        print('\n','%s is connected' % 'MySQL fund db','\n')
        # ------------------------------------------------------------------------------------------------------------------------
        def spreadSheetFileListInDirectory(self, filename_header):
            #요약 : 특정 디렉터리에 특정 파일 이름을 가진 스프레드시트 파일을 목록화함.
            #input: 파일명의 접두어(str), 대상이 될 확장자 목록(list), 파일이 저장된 경로(Fund_settings.fundXlsxFile_path)
            #return: 경로, 파일명, 경로를 포함한 파일명을 묶은 목록(dict)
            #참고 : appendToTable에서 참조함.

            import glob as gb
            from os import path
            with_path = []
            filetypes = ['*.xlsx','*.xls','*.csv']
            for filetype in filetypes:
                temp = gb.glob(Fund_settings.fundXlsxFile_path+filename_header+filetype)
                with_path = with_path + temp
            without_path = [path.basename(file) for file in with_path]
            filelist = {}
            filelist['target_path'] = Fund_settings.fundXlsxFile_path
            filelist['with_path'] = with_path
            filelist['with_path'].sort()
            filelist['without_path'] = without_path
            filelist['without_path'].sort()
            filelist['file_counter'] = len(with_path)
            if len(filelist['with_path']) == len(filelist['without_path']):
                return(filelist)
        def move_to_applied_folder(self, filename_with_path, filename_without_path, target_path):
            #요약 : 특정 폴더안에 있는 파일을 그 하위폴더로 옮기는 함수
            #input: 경로포함 파일명, 경로없는 파일명, 하위폴더명(Fund_settings.fundXlsxFileBackupDirectory_afterDBupdate)
            #return: 없음.
            #참고 : appendToTable에서 참조함.

            import os
            target_folder_name = Fund_settings.fundXlsxFileBackupDirectory_afterDBupdate
            try:
                os.rename(filename_with_path,target_path+'\\'+target_folder_name+'\\'+filename_without_path)
            except:
                print('Exists!! ' + target_path+'\\'+target_folder_name+'\\'+filename_without_path)
        def appendToTable(self, filename_header):
            #요약 : 파일명의 접두어를 기준으로, 접두어와 같은 이름의 함수를 불러내어 데이터를 정제한 후 DB에 업로드하는 함수.
            #input: 파일명의 접두어(str), fund.fund_updateXLSXtoMySQLDB() 상속
            #return: 없음.
            #실행순서1 : spreadSheetFileListInDirectory를 실행하여 DB에 넣을 기준파일들의 목록을 만듦.
            #실행순서2 : 접두어와 일치하는 fund.fund_updateXLSXtoMySQLDB()의 모듈을 불러와서 Dataframe을 리턴받음.
            #실행순서3 : 접두어에 따라 쿼리를 추가적으로 실행한 뒤, Dataframe을 접두어와 같은 테이블로 업로드함.
            #실행순서4 : <declassified> 실행하여 DB 업로드가 끝난 파일을 하위폴더로 옮겨둠.

            class_call1 = fund.fund_updateXLSXtoMySQLDB()
            #파일 헤더를 기준으로 기준 폴더 안의 xlsx/xls 파일의 목록을 불러옴.
            filelist = self.spreadSheetFileListInDirectory(filename_header)
            print(filelist['file_counter'],' of ',filename_header,' header files exists in ',Fund_settings.fundXlsxFile_path)
            print('--  DB update is started : ',filename_header,'-'*17)
            for i in range(0,filelist['file_counter']):
                try:
                    #class_call1으로 상속받은 클래스의 하위 함수들의 이름 중, 파일 헤더와 매칭되는 함수를 func_call로 불러옴.(파일 헤더에 맞는 dataframe 모양 조정 함수들)
                    func_call = getattr(class_call1,filename_header)
                    #불러온 함수 실행하여 정제된 dataframe을 data로 return 받음.
                    data = func_call(filelist['with_path'][i],filelist['without_path'][i])
                    #테이블 전체 갱신이 필요한 codematch, fas3421의 경우, 전체 테이블을 delete 하고 insert into 함.
                    if filename_header in Fund_settings.fundTableReplace :
                        self.cs.execute('delete from '+filename_header+';')
                        self.con.commit()
                        print('fund.',filename_header,' is deleted.')
                    #파일 헤더를 기준으로 조정된 데이터프레임의 내용을, <declassified> DB 아래의 헤더 이름의 테이블에 추가함.
                    data.to_sql(con=self.conn, name=filename_header, if_exists='append', index=False)
                    #기준 파일을 다른 폴더로 옮기기
                    self.move_to_applied_folder(filelist['with_path'][i],filelist['without_path'][i],filelist['target_path'])
                    print(filelist['without_path'][i],' is inserted')
                except PermissionError:
                    #기준 파일이 열려있으면 파일 접근 에러가 생김.
                    print('A file is opened. Retry after closing this file.','-'*20,'\n')
                except IndexError:
                    print('Additional fund might be existed. Check the variable fundcode_list in <declassified> mothod.','-'*20,'\n')
            print('--  DB update is done : ',filename_header,'-'*20,'\n')
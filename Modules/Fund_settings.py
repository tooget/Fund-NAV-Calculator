#-*- coding:utf-8 -*-

#<declassified>

def codeAvailable(date):
    #tempDict = {}
    #[tempDict.update(self.<declassified>[x]) for x in self.<declassified>.keys() if self.<declassified>[x] is not None]
    #codeAvailable = [ x for x in tempDict if tempDict[x] is True ] + list(self.<declassified>.keys())
    #codeAvailable.sort
    import pymysql
    pymysql.install_as_MySQLdb()
    import MySQLdb
    #DB Charset이 utf8, Collation Name이 utf8_general_ci 인지 확인할 것!
    #pandas Dataframe.to_sql() 메소드가 pymysql의 connection으로는 연결되지 않음.
    con = pymysql.connect(
            host=fundDB_host,
            port=fundDB_port,
            user=fundDB_user,
            passwd=fundDB_passwd,
            db=fundDB_name,
            charset=fundDB_charset,
            )  
    cs = con.cursor(MySQLdb.cursors.DictCursor)
    cs.execute("select <declassified> from <declassified> where <declassified> <= '" + date + "';")
    result_set = cs.fetchall()
    for x in result_set:
        x['<declassified>'] = '<declassified>' + str(x['<declassified>'])
    return(result_set)
def codeofMotherFunds_SpecificHeaderOnly():
    import pymysql
    pymysql.install_as_MySQLdb()
    import MySQLdb
    #DB Charset이 utf8, Collation Name이 utf8_general_ci 인지 확인할 것!
    #pandas Dataframe.to_sql() 메소드가 pymysql의 connection으로는 연결되지 않음.
    con = pymysql.connect(
            host=fundDB_host,
            port=fundDB_port,
            user=fundDB_user,
            passwd=fundDB_passwd,
            db=fundDB_name,
            charset=fundDB_charset,
            )  
    cs = con.cursor(MySQLdb.cursors.DictCursor)
    #굳이 설립일 정보를 쓰지 않음
    cs.execute("select <declassified> from <declassified> where <declassified> not in ('<declassified>') and <declassified> like '<declassified>';")
    result_set = cs.fetchall()
    for x in result_set:
        x['<declassified>'] = '<declassified>' + str(x['<declassified>'])
    return(result_set)
def codePairs_mothernchild(date):
    import pymysql
    pymysql.install_as_MySQLdb()
    import MySQLdb
    #DB Charset이 utf8, Collation Name이 utf8_general_ci 인지 확인할 것!
    #pandas Dataframe.to_sql() 메소드가 pymysql의 connection으로는 연결되지 않음.
    con = pymysql.connect(
            host=fundDB_host,
            port=fundDB_port,
            user=fundDB_user,
            passwd=fundDB_passwd,
            db=fundDB_name,
            charset=fundDB_charset,
            )  
    cs = con.cursor(MySQLdb.cursors.DictCursor)
    cs.execute("select <declassified>, <declassified> from <declassified> where <declassified> = '<declassified>' and <declassified> <= '" + date + "';")
    result_set = cs.fetchall()
    for x in result_set:
        x['<declassified>'] = '<declassified>' + str(x['<declassified>'])
        x['<declassified>'] = '<declassified>' + str(x['<declassified>'])





    return(result_set)
def codePairs_children(date):
    ref = codePairs_mothernchild(date)
    motherCodes_HavingChildren = list( set( [pair['<declassified>'] for pair in ref] ) )

    codePairs_Children = {}
    for motherCode in motherCodes_HavingChildren:
        temp = [x['<declassified>'] for x in filter( lambda x: x['<declassified>'] == motherCode, ref) ]
        codePairs_Children[motherCode] = temp
        del(temp)
    return(codePairs_Children)


#fund_extractXLSXfromFAS.py
fasExePath = "<declassified>"
fasLoginId = '<declassified>'
fasLoginPw = '<declassified>'
fasExtractTRList = ['<declassified>','<declassified>','<declassified>','<declassified>','<declassified>','<declassified>','<declassified>']

fundXlsxFile_path = '<declassified>'
fundXlsxFileBackupDirectory_afterDBupdate = '<declassified>'
fasGuiImagePath = '<declassified>\\*.png'

fundDB_host='<declassified>'
fundDB_port='<declassified>'
fundDB_user='<declassified>'
fundDB_passwd='<declassified>'
fundDB_name='<declassified>'
fundDB_charset='utf8'
fundTableReplace = ['<declassified>','<declassified>']

fundQueryPath = '<declassified>'
iNAVqueryFilename = '<declassified>'
DeriPnLqueryFilename = '<declassified>'
iNAVclassfundQueryFilename = '<declassified>'

bizmekaEdmGuiImagePath = '<declassified>\\*.png'
bizmekaAccountGuiImagePath = '<declassified>\\*.png'


# 분개장 파일 참조 경로 --------------------------------------------------------
accountLedgerFilePath = r'<declassified>'
# 증빙 폴더 경로 ------------------------------------------------------------
accountLedgerDirectoryPath = r'<declassified>'
# 증빙 폴더 경로 ------------------------------------------------------------
accountHtsXlsxFilePath = r'<declassified>'
# 규정 폴더 경로 ------------------------------------------------------------
# regulation_dir = r'<declassified>'

accountDB_host='<declassified>'
accountDB_port='<declassified>'
accountDB_user='<declassified>'
accountDB_passwd='<declassified>'
accountDB_name='<declassified>'
accountDB_charset='utf8'
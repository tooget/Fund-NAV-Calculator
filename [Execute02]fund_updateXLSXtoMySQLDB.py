#-*- coding:utf-8 -*-

#1.기본경로 내에 <declassified>.xlsx의 파일을 처리한다.
#2.기본경로 하위에 files_applied_to_db 폴더가 있어야, 파일을 처리하고 이동시킨다. file_struc 클래스 참조
#3.파일명은 '_'로 구분하며, 첫번째는 filename_headers 요소값, Arrange_for_db의 함수명 앞부분과 일치해야 함.

from Modules import Fund_modules
from Modules import Fund_settings


# 실행 -------------------------------------------------------------------------------------------------------------------
updateXLSX = Fund_modules.fund.fund_updateXLSXtoMySQLDB_interface()
filename_headers = [value for value in dir(Fund_modules.fund.fund_updateXLSXtoMySQLDB) if value.find('__') is -1]
print('-'*50,'\n',Fund_settings.fundXlsxFile_path,'\n',filename_headers,'\n','-'*50,'\n\n')
for filename_header in filename_headers:
    #print(filename_header)
    updateXLSX.appendToTable(filename_header)
# ------------------------------------------------------------------------------------------------------------------------
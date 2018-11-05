#-*- coding:utf-8 -*-

from Modules import Fund_settings
from Modules import Fund_modules
import time


fas_gui = Fund_modules.fund.fund_extractXLSXfromFAS()

fasLoginID = Fund_settings.fasLoginId
fasLoginPW = Fund_settings.fasLoginPw
fas_ExePath = Fund_settings.fasExePath
fasGuiImage_path = Fund_modules.general_func.filesInDirectory(Fund_settings.fasGuiImagePath)

#fromNdaysBefore = int(input("Export xlsx files from N days before : "))
dateToExtract = str(input("Input Date to export xlsx files : "))
dateToExtract = dateToExtract.replace('-','')

#datelist = Fund_modules.general_func.create_datelist(fromNdaysBefore,'filename')
datelist = [dateToExtract]
fundcode = Fund_settings.codeAvailable(datelist[0])
fas_tr_list = Fund_settings.fasExtractTRList

print("-"*50)
print("Start date : ",datelist[0])
print("End Date : ",datelist[-1])
print("Target Funds : ",fundcode)
print("Target FAS TR Numbers : ",fas_tr_list)
print("-"*50)

startTimeStamp = time.time()
print("START : ",time.ctime(startTimeStamp))
print("-"*50)

from threading import Thread
from subprocess import run as srun
class fas_execute(Thread):
    def __init__(self,exepath):
        Thread.__init__(self)
        self.exepath = exepath
    def run(self):
        srun(self.exepath)

t1 = fas_execute(fas_ExePath)
t1.start()
print('FAS executed.')


fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[0],'exist',10)
fas_gui.imageRightSideClick(fasGuiImage_path[0])
fas_gui.pasteKey(fasLoginID)
fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[1],'exist',10)
fas_gui.imageRightSideClick(fasGuiImage_path[1])
fas_gui.pasteKey(fasLoginPW)
fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[2],'exist',10)
fas_gui.imageCenterClick(fasGuiImage_path[2])
fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[3],'exist',10)
fas_gui.imageCenterDoubleClick(fasGuiImage_path[4])

saveComplete = []

for fas_tr_num in fas_tr_list:

    try:
        fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[5],'exist',5)
        fas_gui.imageCenterClick(fasGuiImage_path[5])
    except TypeError:
        print('PASS : Find another TR Empty inputbox.')
        fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[15],'exist',5)
        fas_gui.imageCenterClick(fasGuiImage_path[15])
        
    time.sleep(0.3)
    fas_gui.pasteKey(fas_tr_num)
    fas_gui.inputKeys('enter')

    #<declassified>
    if fas_tr_num in ['<declassified>']:
        for date in datelist[0:1]:

            print('START: ',fas_tr_num,date)

            # 04 조회 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[25],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[25])
            # 05 처리중 돌아가는 이미지가 없는지 확인
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[9],'notexist',10)

            # 07 엑셀 추출 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[21],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[21])
            # 08 엑셀 X 버튼 클릭
            # 엑셀 창 뜨는 게 유난히 오래 걸림.
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[11],'exist',25)
            fas_gui.imageCenterClick(fasGuiImage_path[11])
            # 09 엑셀 종료 확인창 팝업시 저장 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[12],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[12])
            # 10 다른이름으로 저장 아이콘 이미지가 뜨는지 확인
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[13],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[13])
            # 11 경로 및 파일이름 붙여넣기
            fileNamenPath = Fund_settings.fundXlsxFile_path+'<declassified>'+fas_tr_num+'_'+date+'.xlsx'
            #fileNamenPath = Fund_settings.fundXlsxFile_path+'fas'+fas_tr_num+'_'+date+'.xlsx'
            fas_gui.pasteKey(fileNamenPath)
            saveComplete.append(fileNamenPath)
            # 12 저장 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[14],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[14])
            print('COMPLETE: <declassified>',saveComplete[-1])
        time.sleep(1)
    #<declassified>
    elif fas_tr_num in ['<declassified>']:
        for code in fundcode:
            for date in datelist:

                print('START: ',fas_tr_num,code['<declassified>'],date)

                # 01 달력 아이콘 누르고 엔터
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[6],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[6])
                time.sleep(0.3)
                fas_gui.inputKeys('enter')
                # 02 날짜 순서대로 입력하고 엔터
                for keys in date:
                    fas_gui.inputKeys(keys)
                    time.sleep(0.15)
                fas_gui.inputKeys('enter')
                # 03 펀드 입력부에 펀드코드를 붙여넣기
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[17],'exist',10)
                fas_gui.imageRightSideClick(fasGuiImage_path[17])
                fas_gui.inputHotkeys('ctrl','a')
                fas_gui.pasteKey(code['<declassified>'][1:])
                # 04 조회 버튼 클릭
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[20],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[20])
                # 05 처리중 돌아가는 이미지가 없는지 확인
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[9],'notexist',10)
                try:
                    # 06 자료 없음 메세지가 있는지 확인 후 확인 버튼 클릭 > 다음으로 넘어감
                    fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[18],'exist',5)
                    fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[19],'exist',1)
                    fas_gui.imageCenterClick(fasGuiImage_path[19])
                    print('PASS: Skip ', fas_tr_num, code['<declassified>'], date)
                except TypeError:
                    # 06 자료 없음 메세지가 있는지 확인 후 없으면 > 계속 진행
                    # 07 엑셀 추출 버튼 클릭
                    fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[21],'exist',10)
                    fas_gui.imageCenterClick(fasGuiImage_path[21])
                    # 08 엑셀 X 버튼 클릭
                    fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[11],'exist',10)
                    fas_gui.imageCenterClick(fasGuiImage_path[11])
                    # 09 엑셀 종료 확인창 팝업시 저장 버튼 클릭
                    fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[12],'exist',10)
                    fas_gui.imageCenterClick(fasGuiImage_path[12])
                    # 10 다른이름으로 저장 아이콘 이미지가 뜨는지 확인
                    fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[13],'exist',10)
                    fas_gui.imageCenterClick(fasGuiImage_path[13])
                    # 11 경로 및 파일이름 붙여넣기
                    fileNamenPath = Fund_settings.fundXlsxFile_path+'<declassified>'+fas_tr_num+'_'+code['<declassified>']+'_'+date+'.xlsx'
                    #fileNamenPath = Fund_settings.fundXlsxFile_path+'fas'+fas_tr_num+'_'+date+'.xlsx'
                    fas_gui.pasteKey(fileNamenPath)
                    saveComplete.append(fileNamenPath)
                    # 12 저장 버튼 클릭
                    fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[14],'exist',10)
                    fas_gui.imageCenterClick(fasGuiImage_path[14])
                    print('COMPLETE: <declassified>',saveComplete[-1])
                time.sleep(1)
    #<declassified>
    elif fas_tr_num in ['<declassified>']:
        for date in datelist:

            print('START: ',fas_tr_num,date)

            # 01 기준일자(시작) 아이콘 누르고 엔터
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[22],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[22])
            time.sleep(0.3)
            for keys in date:
                fas_gui.inputKeys(keys)
                time.sleep(0.15)
            fas_gui.inputKeys('enter')
            # 02 기준일자(끝) 아이콘 누르고 엔터
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[23],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[23])
            time.sleep(0.3)
            for keys in date:
                fas_gui.inputKeys(keys)
                time.sleep(0.15)
            fas_gui.inputKeys('enter')
            # 03 조회 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[8],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[8])
            # 04 처리중 돌아가는 이미지가 없는지 확인
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[9],'notexist',10)
            # 05 엑셀 추출 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[24],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[24])
            # 06 엑셀 X 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[11],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[11])
            # 07 엑셀 종료 확인창 팝업시 저장 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[12],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[12])
            # 08 다른이름으로 저장 아이콘 이미지가 뜨는지 확인
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[13],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[13])
            # 09 경로 및 파일이름 붙여넣기
            fileNamenPath = Fund_settings.fundXlsxFile_path+'<declassified>'+fas_tr_num+'_'+date+'.xlsx'
            #fileNamenPath = Fund_settings.fundXlsxFile_path+'<declassified>'+fas_tr_num+'_'+date+'.xlsx'
            fas_gui.pasteKey(fileNamenPath)
            saveComplete.append(fileNamenPath)
            # 12 저장 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[14],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[14])
            print('COMPLETE: <declassified>',saveComplete[-1])
            time.sleep(1)
    #<declassified>
    elif fas_tr_num in ['<declassified>']:
        for date in datelist:
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[6],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[6])
            time.sleep(0.3)
            fas_gui.inputKeys('enter')

            for keys in date:
                fas_gui.inputKeys(keys)
                time.sleep(0.15)

            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[8],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[8])
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[9],'notexist',10)
            time.sleep(0.5)
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[10],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[10])
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[11],'exist',15)
            fas_gui.imageCenterClick(fasGuiImage_path[11])
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[12],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[12])
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[13],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[13])
            fileNamenPath = Fund_settings.fundXlsxFile_path+'<declassified>'+fas_tr_num+'_'+date+'.xlsx'
            fas_gui.pasteKey(fileNamenPath)
            saveComplete.append(fileNamenPath)
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[14],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[14])
            print('COMPLETE: <declassified>',saveComplete[-1])
            time.sleep(1)
    #<declassified>
    elif fas_tr_num in ['<declassified>']:
        for code in fundcode:
            for date in datelist:
                print('START: ',fas_tr_num,code['<declassified>'],date)
                time.sleep(0.5)                
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[6],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[6])
                time.sleep(0.5)
                fas_gui.inputKeys('enter')

                for keys in date:
                    fas_gui.inputKeys(keys)
                    time.sleep(0.15)
                fas_gui.inputKeys('enter')

                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[16],'exist',10)
                fas_gui.imageRightSideClick(fasGuiImage_path[16])

                fas_gui.inputHotkeys('ctrl','a')
                fas_gui.pasteKey(code['<declassified>'][1:])

                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[8],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[8])
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[9],'notexist',10)
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[10],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[10])
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[11],'exist',15)
                fas_gui.imageCenterClick(fasGuiImage_path[11])
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[12],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[12])
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[13],'exist',10)
                fileNamenPath = Fund_settings.fundXlsxFile_path+'<declassified>'+fas_tr_num+'_'+code['<declassified>']+'_'+date+'.xlsx'
                #fileNamenPath = Fund_settings.fundXlsxFile_path+'fas'+fas_tr_num+'_'+date+'.xlsx'
                fas_gui.pasteKey(fileNamenPath)
                saveComplete.append(fileNamenPath)
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[14],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[14])
                print('COMPLETE: <declassified>',saveComplete[-1])
                time.sleep(1)
    #<declassified>
    elif fas_tr_num in ['<declassified>']:
        for code in fundcode:
            for date in datelist:
                print('START: ',fas_tr_num,code['<declassified>'],date)
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[6],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[6])
                time.sleep(0.3)
                fas_gui.inputKeys('enter')

                for keys in date:
                    fas_gui.inputKeys(keys)
                    time.sleep(0.15)
                fas_gui.inputKeys('enter')

                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[7],'exist',10)
                fas_gui.imageRightSideClick(fasGuiImage_path[7])

                fas_gui.inputHotkeys('ctrl','a')
                fas_gui.pasteKey(code['<declassified>'][1:])

                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[8],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[8])
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[9],'notexist',10)
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[10],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[10])
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[11],'exist',15)
                fas_gui.imageCenterClick(fasGuiImage_path[11])
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[12],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[12])
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[13],'exist',10)
                fileNamenPath = Fund_settings.fundXlsxFile_path+'<declassified>'+fas_tr_num+'_'+code['<declassified>']+'_'+date+'.xlsx'
                #fileNamenPath = Fund_settings.fundXlsxFile_path+'fas'+fas_tr_num+'_'+date+'.xlsx'
                fas_gui.pasteKey(fileNamenPath)
                saveComplete.append(fileNamenPath)
                fas_gui.imageCenterClick(fasGuiImage_path[14])
                print('COMPLETE: <declassified>',saveComplete[-1])
                time.sleep(1)
    #<declassified>
    elif fas_tr_num in ['<declassified>']:
        for date in datelist:

            print('START: ',fas_tr_num,date)

            # 01 기준일자(시작) 아이콘 누르고 엔터
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[26],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[26])
            time.sleep(0.3)
            for keys in date:
                fas_gui.inputKeys(keys)
                time.sleep(0.15)
            fas_gui.inputKeys('enter')
            # 02 기준일자(끝) 아이콘 누르고 엔터
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[27],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[27])
            time.sleep(0.3)
            for keys in date:
                fas_gui.inputKeys(keys)
                time.sleep(0.15)
            fas_gui.inputKeys('enter')
            # 03 조회 버튼 클릭
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[28],'exist',10)
            fas_gui.imageCenterClick(fasGuiImage_path[28])
            # 04 처리중 돌아가는 이미지가 없는지 확인
            fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[9],'notexist',10)

            try:
                # 06 자료 없음 메세지가 있는지 확인 후 없으면 > 계속 진행
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[18],'notexist',5)
                # 07 엑셀 추출 버튼 클릭
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[10],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[10])
                # 08 엑셀 X 버튼 클릭
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[11],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[11])
                # 09 엑셀 종료 확인창 팝업시 저장 버튼 클릭
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[12],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[12])
                # 10 다른이름으로 저장 아이콘 이미지가 뜨는지 확인
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[13],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[13])
                # 11 경로 및 파일이름 붙여넣기
                fileNamenPath = Fund_settings.fundXlsxFile_path+'<declassified>'+fas_tr_num+'_'+date+'.xlsx'
                #fileNamenPath = Fund_settings.fundXlsxFile_path+'fas'+fas_tr_num+'_'+date+'.xlsx'
                fas_gui.pasteKey(fileNamenPath)
                saveComplete.append(fileNamenPath)
                # 12 저장 버튼 클릭
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[14],'exist',10)
                fas_gui.imageCenterClick(fasGuiImage_path[14])
                print('COMPLETE: <declassified>',saveComplete[-1])
            except TypeError:
                # 06 자료 없음 메세지가 있는지 확인 후 확인 버튼 클릭 > 다음으로 넘어감
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[18],'exist',5)
                fas_gui.checkImagesOnscreenLoop(fasGuiImage_path[19],'exist',1)
                fas_gui.imageCenterClick(fasGuiImage_path[19])
                print('PASS: Skip ', fas_tr_num, date)
            time.sleep(1)




#fas_gui.inputHotkeys('alt','f4')
[print(x, saveComplete[x]) for x in range(len(saveComplete))]

endTimeStamp = time.time()
print("-"*50)
print("END : ",time.ctime(endTimeStamp))
print("Running Time : ",round(endTimeStamp-startTimeStamp,2))
print("-"*50)
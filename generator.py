# [Flowinus]

import os
import csv
import xlwings as xw
import shutil
from datetime import date

xl_file     =   'VW_AQ_EOP_SWUTR_CR.xlsx'
path        =   'csvFolder/'         #csv 파일 모음폴더
Tester      =   ''
fileName    =   ''
TCnumber    =   ''

personal    =   []
file_list   =   os.listdir(path)
csv_list    =   []            #csv 파일 리스트
testCaseID  =   []            #TC ID
note        =   []            #csv 파일 내용 문자열
functionName=   []            #test 이름
functionNum =   []            #test 개수
SWDDS       =   []            #SWDDS note
swddsCode   =   []            #test의 SWDDS코드
testName    =   []

print('파일 읽어오는중..')
txt = open('properties.txt', 'r',encoding='utf-8-sig')

for i in txt:
    personal.append(i.split(':'))
Tester = personal[0][1].strip('\n')
fileName = personal[2][1].strip('\n')
if personal[1][1] != '\n':
    date = personal[1][1].strip('\n')
else:
    date = str(date.today())
TCnumber = personal[3][1].strip('\n')
txt.close()

if len(file_list) == 0:
    print("There is no file\n")
else:
    for i in range(len(file_list)):
        if file_list[i].find('.csv') > 0:
            csv_list.append(file_list[i].replace('.csv',''))
    for i in range(len(csv_list)):
        data = csv.reader(open(str(path+csv_list[i]+'.csv'),encoding='cp949'))
        for j in data:
            note.append(j)
        name = note[1][0].strip('test name:').split('_test')
        print(name)
        testName.append(name[0])
        note = []

    name = testName[0]
    countNum = 1
    #테스트 개수파악
    for i in range(1,len(testName)):
        if i == len(testName)-1:
            functionName.append(name)
            functionNum.append(countNum+1)
            break
        elif name == testName[i]:
            countNum = countNum + 1
        else:
            functionName.append(name)
            functionNum.append(countNum)
            name = testName[i]
            countNum = 1
    print(str(len(testName))+'개의 파일 발견')
    #파일 읽기
    print("SWDDS 파일내용수집")
    txt = open('resource/VW_AQ_EOP_12_SWE_Design_VW_AQ_EOP_SWDDS.txt', 'r',encoding='utf-8-sig')
    for i in txt:
        SWDDS.append(i)
    print("내용찾는중")
    #이름에 대한 SWDDS코드 찾기
    for i in range(len(functionName)):
        for j in range(len(SWDDS)):
            if SWDDS[j].find(functionName[i]) > 0:
                if SWDDS[j].find('SWDDS'):
                    print(SWDDS[j])
                    data = SWDDS[j].split(' ')
                    if data[1][0] == '[':
                        swddsCode.append(data[1])
                    else:
                        swddsCode.append(str('['+data[2]+']'))
                    break
    for i in range(len(swddsCode)):
        swddsCode[i] = swddsCode[i].strip('[SWDDS.').strip(']')
    txt.close()
    #print("\n")
    #print(functionName)
    #print(functionNum)
    #print(swddsCode)

    #엑셀자동생성
    file_list = os.listdir()
    for i in range(len(file_list)):
        if file_list[i].find(xl_file.strip('.xlsx') + '_' + TCnumber + '.xlsx') != -1:
            os.remove(xl_file.strip('.xlsx') + '_' + TCnumber + '.xlsx')
    print("빈 엑셀 생성")
    shutil.copy(str('resource/'+xl_file), xl_file) # 결과 출력을 위한 빈파일 생성
    os.rename(xl_file, xl_file.strip('.xlsx') + '_' + TCnumber +'.xlsx')
    xlbook = xw.Book(xl_file.strip('.xlsx') + '_' + TCnumber +'.xlsx')
    sheet = xlbook.sheets['filename.c']
    Ycell = 0
    for i in range(len(functionName)):
        if functionNum[i] == 1:
            sheet.range('C' + str(5 + Ycell)).value = str('SWDDS.' + swddsCode[i])      # SWDDS출력
            sheet.range('B' + str(5 + Ycell)).value = str('SWUTS-F.' + swddsCode[i])    # SWUTS출력
            sheet.range('E' + str(5 + Ycell)).value = functionName[i]                   # unit 이름출력
            sheet.range('F' + str(5 + Ycell)).value = Tester                            # 테스터 출력
            sheet.range('Z' + str(5 + Ycell)).value = date                              # 날짜
            Ycell = Ycell + 1
        else:
            for j in range(functionNum[i]):
                sheet.range('C' + str(5 + Ycell)).value = str('SWDDS.' + swddsCode[i]+'_')+str(j+1)     # SWDDS출력
                sheet.range('B' + str(5 + Ycell)).value = str('SWUTS-F.' + swddsCode[i]+'_')+str(j+1)   # SWUTS출력
                sheet.range('E' + str(5 + Ycell)).value = functionName[i]                               # unit 이름출력
                sheet.range('F' + str(5 + Ycell)).value = Tester                                        # 테스터 출력
                sheet.range('Z' + str(5 + Ycell)).value = date                                          # 날짜
                Ycell = Ycell + 1
    xlbook.sheets['filename.c'].name = fileName
    print(functionName)
    print(functionNum)
    print(swddsCode)
    print("생성완료")
    xlbook.save()
    xlbook.close()
    print("end")


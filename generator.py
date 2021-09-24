# [Flowinus]

import os
import csv
import win32com.client
import xlwings as xw
import shutil
from datetime import date

print("\n┌───────────────────────────────────────────┐")
print("│         Unit TC report generator          │")
print("└───────────────────────────────────────────┘\n")
print("엑셀에 기입하려고 하는 CSV 파일은 csvFolder 에 넣고")
print("테스트보고서와 Test_Result폴더는 testReport 폴더넣어주세요")
print("readme.txt 를 수정하시면 세부사항을 기입할 수 있습니다.\n")
input("진행할려면 엔터를 눌러주세요...\n\n")
print("FLOWINUS.")

xl_file     =   'VW_AQ_EOP_SWUTR_CR.xlsx'   #CT 빈파일
path        =   'csvFolder/'                #csv 파일 모음폴더
Tester      =   ''                          #테스터
fileName    =   ''                          #파일이름
TCnumber    =   ''                          #테스트넘버
TCresultName=   ''                          #

personal    =   []            #readme파일 읽기위한 빈통
file_list   =   os.listdir(path)
csv_list    =   []            #csv 파일 리스트
testCaseID  =   []            #TC ID
note        =   []            #csv 파일 내용 문자열
functionName=   []            #test 이름
functionNum =   []            #test 개수
SWDDS       =   []            #SWDDS note
swddsCode   =   []            #test의 SWDDS코드
testName    =   []            #
testResult  =   []            #

bindata     =   []            #빈통

#커버리지
StateCoverage  =    []
SCNPercent     =    []
SCNTest        =    []
SCNTotal       =    []
BrechCoverage  =    []
BCNPercent     =    []
BCNTest        =    []
BCNTotal       =    []

print('파일 읽어오는중..')
txt = open('readme.txt', 'r',encoding='utf-8-sig')
excelApp1 = win32com.client.dynamic.Dispatch('Excel.Application')
excelApp1.Quit()
for i in txt:
    personal.append(i.split(':'))
Tester = personal[0][1].strip('\n')
fileName = personal[2][1].strip('\n')
if personal[1][1] != '\n':
    date = personal[1][1].strip('\n')
else:
    date = str(date.today())
TCnumber = personal[3][1].strip('\n').strip(' ')
TCresultName = str('test_' + personal[4][1].strip('\n').replace(' ','') + '.xlsx')
txt.close()

if len(file_list) == 0:
    print("There is no file\n")
else:
    for i in file_list:
        print(i)
    for i in range(len(file_list)):
        if file_list[i].find('.csv') > 0:
            csv_list.append(file_list[i].replace('.csv',''))
    for i in range(len(csv_list)):
        data = csv.reader(open(str(path+csv_list[i]+'.csv'),encoding='cp949'))
        for j in data:
            note.append(j)
        name = note[1][0].strip('test name:').split('_test')
        testName.append(name[0])
        note = []

    name = testName[0]
    countNum = 1
    #테스트 개수파악
    for i in range(1,len(testName)+1):
        if i == len(testName):
            functionName.append(name)
            functionNum.append(countNum)
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
    print("SWDDS 탐색중")
    #이름에 대한 SWDDS코드 찾기
    for i in range(len(functionName)):
        for j in range(len(SWDDS)):
            if SWDDS[j].find(functionName[i]) > 0:
                if SWDDS[j].find('SWDDS'):
                    data = SWDDS[j].split(' ')
                    if data[1][0] == '[':
                        swddsCode.append(data[1])
                    else:
                        swddsCode.append(str('['+data[2]+']'))
                    break
    for i in range(len(swddsCode)):
        swddsCode[i] = swddsCode[i].strip('[SWDDS.').strip(']')
    txt.close()
    print("테스트 보고서 유무확인")
    file_list = os.listdir('testReport/')
    if len(file_list) > 0:
        for i in range(len(file_list)):
            if file_list[i] == TCresultName:
                fileOnOff = 1
                print("발견")
                resultBook = xw.Book(str('testReport/'+TCresultName))
                sheet = resultBook.sheets['Sheet0']
                i = 1
                data = str(sheet[str('A1')].value)
                print("탐색중.....",end='')
                while '함수별 커버리지' != data:

                    i = i + 1
                    data = str(sheet['A' + str(i)].value)
                    if i > 5000:
                        break
                print('')
                while 0 > str(sheet['B' + str(i)].value).find('총 함수 개수'):
                    i = i + 1
                    testResult.append(str(sheet['B' + str(i)].value))
                    testResult.append(str(sheet['C' + str(i)].value))
                    testResult.append(str(sheet['D' + str(i)].value))

                for i in range(len(functionName)):
                    for j in range(len(testResult)):
                        if functionName[i] == testResult[j]:
                            StateCoverage.append(testResult[j + 1])
                            BrechCoverage.append(testResult[j + 2])
                for i in range(len(functionName)):
                    if StateCoverage[i] != 'N/A':
                        bindata = StateCoverage[i].split('%')
                        SCNPercent.append(float(bindata[0]))
                        bindata = bindata[1].split('/')
                        SCNTest.append(float(bindata[0].strip('(')))
                        SCNTotal.append(float(bindata[1].strip(')')))
                    else:
                        SCNPercent.append('N/A')
                        SCNTest.append('N/A')
                        SCNTotal.append('N/A')
                    if BrechCoverage[i] != 'N/A':
                        bindata = BrechCoverage[i].split('%')
                        BCNPercent.append(float(bindata[0]))
                        bindata = bindata[1].split('/')
                        BCNTest.append(float(bindata[0].strip('(')))
                        BCNTotal.append(float(bindata[1].strip(')')))
                    else:
                        BCNPercent.append('N/A')
                        BCNTest.append('N/A')
                        BCNTotal.append('N/A')
                print('완료')
                #resultBook.close()
                resultBook.app.quit()
            else:
                fileOnOff = 0
                print("매칭에러")
    else:
        fileOnOff = 0
        print('파일없음')

    excelApp1 = win32com.client.dynamic.Dispatch('Excel.Application')
    excelApp1.Quit()
    #엑셀자동생성
    file_list = os.listdir()
    for i in range(len(file_list)):
        if file_list[i].find(xl_file.strip('.xlsx') + '-00' + TCnumber + '.xlsx') != -1:
            os.remove(xl_file.strip('.xlsx') + '-00' + TCnumber + '.xlsx')
    print("엑셀 생성")
    shutil.copy(str('resource/'+xl_file), xl_file) # 결과 출력을 위한 빈파일 생성
    os.rename(xl_file, xl_file.strip('.xlsx') + '-00' + TCnumber +'.xlsx')
    xlbook = xw.Book(xl_file.strip('.xlsx') + '-00' + TCnumber +'.xlsx')
    sheet = xlbook.sheets['filename.c']
    Ycell = 0
    print('출력')
    for i in range(len(functionName)):
        if functionNum[i] == 1:
            sheet.range('C' + str(5 + Ycell)).value = str('SWDDS.' + swddsCode[i])      # SWDDS출력
            sheet.range('B' + str(5 + Ycell)).value = str('SWUTS-F.' + swddsCode[i])    # SWUTS출력
            sheet.range('E' + str(5 + Ycell)).value = functionName[i]                   # unit 이름출력
            sheet.range('F' + str(5 + Ycell)).value = Tester                            # 테스터 출력
            sheet.range('Z' + str(5 + Ycell)).value = date                              # 날짜
            sheet.range('D' + str(5 + Ycell)).value = fileName                          # 파일이름
            if fileOnOff == 1:
                if SCNPercent[i] == 'N/A':
                    sheet.range('R' + str(5 + Ycell)).value = str(SCNPercent[i])
                else:
                    sheet.range('R' + str(5 + Ycell)).value = str(SCNPercent[i]) + '%'      # 커버리지 state %
                sheet.range('S' + str(5 + Ycell)).value = SCNTest[i]                        # 커버리지 num test
                sheet.range('T' + str(5 + Ycell)).value = SCNTotal[i]                       # 커버리지 num total
                if BCNPercent[i] =='N/A':
                    sheet.range('V' + str(5 + Ycell)).value = str(BCNPercent[i])
                else:
                    sheet.range('V' + str(5 + Ycell)).value = str(BCNPercent[i]) + '%'      # 커버리지 brench %
                sheet.range('W' + str(5 + Ycell)).value = BCNTest[i]                        # 커버리지 num test
                sheet.range('X' + str(5 + Ycell)).value = BCNTotal[i]                       # 커버리지 num total
            sheet.range('Q' + str(5 + Ycell)).value = '-'                               # -
            sheet.range('H' + str(5 + Ycell)).value = '-'                               # -
            Ycell = Ycell + 1
        else:
            for j in range(functionNum[i]):
                sheet.range('C' + str(5 + Ycell)).value = str('SWDDS.' + swddsCode[i]+'_')+str(j+1)     # SWDDS출력
                sheet.range('B' + str(5 + Ycell)).value = str('SWUTS-F.' + swddsCode[i]+'_')+str(j+1)   # SWUTS출력
                sheet.range('E' + str(5 + Ycell)).value = functionName[i]                               # unit 이름출력
                sheet.range('F' + str(5 + Ycell)).value = Tester                                        # 테스터 출력
                sheet.range('Z' + str(5 + Ycell)).value = date                                          # 날짜
                sheet.range('D' + str(5 + Ycell)).value = fileName                                      # 파일이름
                if fileOnOff == 1:
                    if SCNPercent[i] == 'N/A':
                        sheet.range('R' + str(5 + Ycell)).value = str(SCNPercent[i])
                    else:
                        sheet.range('R' + str(5 + Ycell)).value = str(SCNPercent[i]) + '%'                  # 커버리지 state %
                    sheet.range('S' + str(5 + Ycell)).value = SCNTest[i]                                    # 커버리지 num test
                    sheet.range('T' + str(5 + Ycell)).value = SCNTotal[i]                                   # 커버리지 num total
                    if BCNPercent[i] == 'N/A':
                        sheet.range('V' + str(5 + Ycell)).value = str(BCNPercent[i])
                    else:
                        sheet.range('V' + str(5 + Ycell)).value = str(BCNPercent[i]) + '%'                  # 커버리지 brench %
                    sheet.range('W' + str(5 + Ycell)).value = BCNTest[i]                                    # 커버리지 num test
                    sheet.range('X' + str(5 + Ycell)).value = BCNTotal[i]                                   # 커버리지 num total
                sheet.range('Q' + str(5 + Ycell)).value = '-'                                               # -
                sheet.range('H' + str(5 + Ycell)).value = '-'                                               # -
                Ycell = Ycell + 1
    xlbook.sheets['filename.c'].name = fileName
    print("생성완료")
    xlbook.save()
    xlbook.app.quit()
    char = input("Test Result내 파일 명을 변경하시겠습니까? y/n :")
    if char == 'y':
        file_list = os.listdir('testReport/Test_Result/')
        print("기존파일제거")
        for i in range(len(file_list)):
            if file_list[i].find('SWUTS') >= 0:
                os.remove('testReport/Test_Result/' + file_list[i])
                print('remove:'+file_list[i])
        print("Test Result 파일명 변경중.")
        file_list = os.listdir('testReport/Test_Result/')
        for i in range(len(functionName)):
            for j in range(len(file_list)):
                for k in range(int(functionNum[i])):
                    #print(str(functionName[i] + '_test' + str(k)+'.xls'))
                    if file_list[j] == str(functionName[i]+'_test'+str(k)+'.xls'):
                        if functionNum[i] == 1:
                            os.rename('testReport/Test_Result/'+functionName[i]+'_test'+str(k)+'.xls','testReport/Test_Result/SWUTS-F.'+swddsCode[i]+'.xls')
                        else:
                            os.rename('testReport/Test_Result/' + functionName[i] + '_test' + str(k) + '.xls','testReport/Test_Result/SWUTS-F.' + swddsCode[i] + '_' + str(k + 1) + '.xls')
        char = input("이름이 변경되지 않은 파일을 제거하시겠습니까? y/n :");
        if char == 'y':
            file_list = os.listdir('testReport/Test_Result/')
            print("제거중.")
            for i in range(len(file_list)):
                if file_list[i].find('SWUTS') < 0:
                    os.remove('testReport/Test_Result/'+file_list[i])
                    print("remove:"+file_list[i])
    print("완료")

input('끝. 엔터를 눌러주세요.')


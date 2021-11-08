# [Flowinus]

import os
import csv
import win32com.client
import xlwings as xw
import shutil
from datetime import date
import time

print("\n┌───────────────────────────────────────────┐")
print("│         Unit TC report generator          │")
print("└───────────────────────────────────────────┘\n")
print("엑셀에 기입하려고 하는 CSV 파일은 csvFolder 에 넣고")
print("테스트보고서와 개별테스트보고서는 테스트보고서 내보내기 디렉토리 경로를")
print("testReport폴더로 지정하고 내보내기 하시면 됩니다")
print("readme.txt 를 수정하시면 세부사항을 기입할 수 있습니다.\n")
input("진행할려면 엔터를 눌러주세요...\n\n")
print("FLOWINUS.")

xl_file     =   'Report.xlsx'   #CT 빈파일
CSVpath     =   'csvFolder/'                #csv 파일 모음폴더
Tester      =   ''                          #테스터
fileName    =   ''                          #파일이름
TCnumber    =   ''                          #테스트넘버
TCresultName=   ''                          #

personal    =   []            #readme파일 읽기위한 빈통
file_list   =   os.listdir('testReport/')
csv_list    =   []            #csv 파일 리스트
note        =   []            #csv 파일 내용 문자열
functionName=   []            #test 이름
functionNum =   []            #test 개수
swddsCode   =   []            #test의 SWDDS코드
functionFileName = []         #테스트의 파일위치
testCaseNum =   []            #테스트케이스의 갯수
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
TCresultName = str('test_' + file_list[0] + '.xlsx')
projectName = file_list[0]
for i in txt:
    personal.append(i.split(':'))
Tester = personal[0][1].strip('\n')
fileName = personal[2][1].strip('\n')
if personal[1][1] != '\n':
    date = personal[1][1].strip('\n')
else:
    date = str(date.today())
TCnumber = personal[3][1].strip('\n').strip(' ')
txt.close()

file_list   =   os.listdir(CSVpath)
if len(file_list) == 0:
    print("There is no file\n")
else:
    for i in file_list:                 #CSV 파일 이름 출력
        print(i)
    for i in range(len(file_list)):     #CSV 파일 확장자 제거
        if file_list[i].find('.csv') > 0:
            csv_list.append(file_list[i].replace('.csv',''))
    for i in range(len(csv_list)):      #CSV 파일 열고 테스트 이름 가져오기
        data = csv.reader(open(str(CSVpath+csv_list[i]+'.csv'),encoding='cp949'))
        for j in data:
            note.append(j)
        name = note[1][0].strip('test name:').split('_test')
        testName.append(name[0])
        note = []

    # 각각의 테스트의 갯수 파악
    name = testName[0]
    countNum = 1
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
    print('테스트 이름:\t',end='')
    print(functionName)
    print('각 테스트당 갯수:\t',end='')
    print(functionNum)
    print(str(len(testName))+'개의 파일 발견')

    # 보고서 파일들에서 SWDDS 코드 읽어오기
    file_list = os.listdir(str('testReport/' + projectName + '/Test_Result/'))
    print(functionName)
    for i in range(len(functionName)):
        num = 0
        for j in file_list:
            name = j.split('_test')
            if name[0] == functionName[i]:
                while name[1] != str(num)+'.xls':
                    num = num + 1
                    if num > 50:
                        break
                print(str('testReport/' + projectName + '/Test_Result/'+ name[0] + '_test'+str(num)+'.xls'))
                book = xw.Book('testReport/' +str(projectName) + '/Test_Result/' + str(name[0])+ '_test'+str(num)+'.xls')
                sheet = xw.sheets['Report']
                testCaseNum.append(int(sheet.range('A11').value))
                text = sheet.range('A1').value
                text = text.strip("테스트 보고서 'SWUTS-F.").strip("'").split("_")
                swddsCode.append(text[0])
                book.app.quit()
                time.sleep(0.5)
                A = num
                num = num + 1
                while num < functionNum[i]+A:
                    print(str('testReport/' + projectName + '/Test_Result/' + name[0] + '_test' + str(num) + '.xls'))
                    book = xw.Book('testReport/' + str(projectName) + '/Test_Result/' + str(name[0]) + '_test' + str(num) + '.xls')
                    sheet = xw.sheets['Report']
                    testCaseNum.append(int(sheet.range('A11').value))
                    book.app.quit()
                    time.sleep(0.5)
                    num = num + 1
                break
    print('SWDDS code:\t',end='')
    print(swddsCode)
    print(testCaseNum)

    #이름에 대한 커버리지 찾기

    print("테스트 보고서 유무확인")
    file_list = os.listdir(str('testReport/'+projectName+'/'))
    if len(file_list) > 0:
        for i in range(len(file_list)):
            if file_list[i] == TCresultName:
                fileOnOff = 1
                print("발견")
                resultBook = xw.Book(str('testReport/'+projectName+'/'+TCresultName))
                sheet = resultBook.sheets['Sheet0']
                i = 1
                data = str(sheet[str('A1')].value)
                print("탐색중.....",end='')
                while '함수별 커버리지' != data:
                    i = i + 1
                    data = str(sheet['A' + str(i)].value)
                    if i > 5000:
                        break
                    elif i%50==0:
                        print('.',end='')
                print('')
                while 0 > str(sheet['B' + str(i)].value).find('총 함수 개수'):
                    i = i + 1
                    testResult.append(str(sheet['A' + str(i)].value))
                    testResult.append(str(sheet['B' + str(i)].value))
                    testResult.append(str(sheet['C' + str(i)].value))
                    testResult.append(str(sheet['D' + str(i)].value))

                for i in range(len(functionName)):
                    for j in range(len(testResult)):
                        if functionName[i] == testResult[j]:
                            text = testResult[j - 1]
                            text = text.split('\\')
                            functionFileName.append(text[len(text)-1])
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
                resultBook.app.quit()
            else:
                fileOnOff = 0
    else:
        fileOnOff = 0
        print('커버리지 보고서 파일없음')
        input()
    #엑셀자동생성
    file_list = os.listdir('Result/')
    for i in file_list:
        if i == xl_file:
            os.remove(str('Result/' + xl_file))
    print("엑셀 생성")
    shutil.copy(str('source/'+ xl_file), str('Result/' + xl_file)) # 결과 출력을 위한 빈파일 생성
    os.rename(str('Result/' + xl_file), str('Result/' + xl_file))
    xlbook = xw.Book(str('Result/' + xl_file))
    print('출력',end='')
    sheet = xlbook.sheets['filename.c']
    Ycell = 0
    for i in range(len(functionName)):
        if functionNum[i] == 1:
            sheet.range('C' + str(5 + Ycell)).value = str('SWDDS.' + swddsCode[i])      # SWDDS출력
            sheet.range('B' + str(5 + Ycell)).value = str('SWUTS-F.' + swddsCode[i])    # SWUTS출력
            sheet.range('E' + str(5 + Ycell)).value = functionName[i]                   # unit 이름출력
            sheet.range('F' + str(5 + Ycell)).value = Tester                            # 테스터 출력
            sheet.range('Z' + str(5 + Ycell)).value = date                              # 날짜
            sheet.range('D' + str(5 + Ycell)).value = functionFileName[i]               # 파일이름
            sheet.range('Q' + str(5 + Ycell)).value = testCaseNum[Ycell]                # 테스트케이스 갯수
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
            sheet.range('H' + str(5 + Ycell)).value = '-'                                   # -
            Ycell = Ycell + 1
            print('.',end='')
        else:
            for j in range(int(functionNum[i])):
                sheet.range('C' + str(5 + Ycell)).value = str('SWDDS.' + swddsCode[i]+'_')+str(j+1)     # SWDDS출력
                sheet.range('B' + str(5 + Ycell)).value = str('SWUTS-F.' + swddsCode[i]+'_')+str(j+1)   # SWUTS출력
                sheet.range('E' + str(5 + Ycell)).value = functionName[i]                               # unit 이름출력
                sheet.range('F' + str(5 + Ycell)).value = Tester                                        # 테스터 출력
                sheet.range('Z' + str(5 + Ycell)).value = date                                          # 날짜
                sheet.range('D' + str(5 + Ycell)).value = functionFileName[i]                           # 파일이름
                sheet.range('Q' + str(5 + Ycell)).value = testCaseNum[Ycell]                            # 테스트케이스 갯수
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
                sheet.range('H' + str(5 + Ycell)).value = '-'                                               # -
                Ycell = Ycell + 1
                print('.', end='')
    xlbook.sheets['filename.c'].name = fileName
    print()
    print("생성완료")
    xlbook.save()
    xlbook.app.quit()
    file_list = os.listdir(str('testReport/'+projectName+'/Test_Result/'))
    print("Test Result xl 파일 가져오는 중.")
    for i in range(len(functionName)):
        k = 0
        p = 0
        while k < int(functionNum[i]):
            for j in file_list:
                if j == str(functionName[i] + '_test' + str(k+p) + '.xls'):
                    if int(functionNum[i]) == 1:
                        shutil.copy(str('testReport/' + projectName + '/Test_Result/' + functionName[i] + '_test' + str(k+p) + '.xls'), str('Result/SWUTS-F.' + swddsCode[i] + '.xls'))
                        print('copy: ' + functionName[i] + '_test' + str(k+p) + '.xls   \t\t-->\t SWUTS-F.' + swddsCode[
                            i] + '.xls')
                    else:
                        shutil.copy(str('testReport/' + projectName + '/Test_Result/' + functionName[i] + '_test' + str(k+p) + '.xls'), str('Result/SWUTS-F.' + swddsCode[i] + '_'+str(k+1)+'.xls'))
                        print('copy: ' + functionName[i] + '_test' + str(k+p) + '.xls   \t\t-->\t SWUTS-F.' + swddsCode[
                            i] + '_' + str(k + 1) + '.xls')
                    k = k + 1
            p = p + 1
    print("완료")

input('작업끝.  엔터를 눌러주세요.')


# [Flowinus]
STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE = -11
STD_ERROR_HANDLE = -12

FOREGROUND_BLACK = 0x00
FOREGROUND_BLUE = 0x01  # text color contains blue.
FOREGROUND_GREEN = 0x02  # text color contains green.
FOREGROUND_RED = 0x04  # text color contains red.
FOREGROUND_INTENSITY = 0x08  # text color is intensified.
BACKGROUND_BLUE = 0x10  # background color contains blue.
BACKGROUND_GREEN = 0x20  # background color contains green.
BACKGROUND_RED = 0x40  # background color contains red.
BACKGROUND_INTENSITY = 0x80  # background color is intensified.

import os
import csv
import win32com.client
import xlwings as xw
import shutil
from datetime import datetime
import time
import ctypes
import re

std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
def set_color(color, handle=std_out_handle):
    bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)
    return bool
    set_color(i)
    print("Hello, world!")
set_color(3)
print()
print("\t┌───────────────────────────────────────────┐")
print("\t│                                           │")
print("\t│       Unit TC report generator            │")
print("\t│       Version 5.0.7                       │")
print("\t│       Last update date 22/01/19           │")
print("\t│                             [ NANI ]      │")
print("\t└───────────────────────────────────────────┘\n")
print()
set_color(15)
print('\x1b[97m')
print("엑셀에 기입하려고 하는 CSV 파일은 csvFolder 에 넣고")
print("테스트보고서와 개별테스트보고서는 테스트보고서 내보내기 디렉토리 경로를")
print("testReport폴더로 지정하고 내보내기 하시면 됩니다")
print("실행결과는 Result 폴더에 생성됩니다.")
excelApp1 = win32com.client.dynamic.Dispatch('Excel.Application')
excelApp1.Quit()
modeAnswer = str(input("자동 분류를 하시겠습니까? Y/N :"))
if modeAnswer =='Y':
    modeAnswer = 'y'
elif modeAnswer == 'N':
    modeAnswer = 'n'
if modeAnswer != 'y' and modeAnswer != 'n':
    modeAnswer = 'n'
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
functionFileName    =   []    #테스트의 파일위치
testCaseNum         =   []    #테스트케이스의 갯수
valueList   =   []            #테스트 변수명의 리스트
stubList    =   []            #테스트이름당 스텁이름
testResultPassOrNot =   []    #테스트이름당 OK 또는 NOK
testType    =   []            #테스트당 타입

# 1 = BND
DescriptionMsBND = "Requirement Based Test among Analysis of boundary value of SWDDS."
# 2 = EQV
DescriptionMsEQV = "Requirement Based Test among Equivalence testing of SWDDS."
# 3 = FIT
DescriptionMsFIT = "Fault Injection Test among Error guessing of SWDDS."
# 4 = STATE
DescriptionMsSTA = "Requirement Based Test among Development of positive of SWDDS."


testName    =   []            #
testResult  =   []            #

caseExplain =   []            #각 테스트당 설명저장

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
#txt = open('readme.txt', 'r',encoding='utf-8-sig')
TCresultName = str('test_' + file_list[0] + '.xlsx')
projectName = file_list[0]
#for i in txt:
#    personal.append(i.split(':'))
#Tester = personal[0][1].strip('\n')
#if personal[1][1] != '\n':
#    date = personal[1][1].strip('\n')
#else:
date = datetime.today().strftime('%Y-%m-%d')
#txt.close()

file_list   =   os.listdir(CSVpath)
if len(file_list) == 0:
    print("파일이 없습니다.\n")
else:
    for i in file_list:                 #CSV 파일 이름 출력
        print(i)
    for i in range(len(file_list)):     #CSV 파일 확장자 제거
        if file_list[i].find('.csv') > 0:
            csv_list.append(file_list[i].replace('.csv',''))
            caseExplain.append("")     #설명 리스트 늘리기
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
            stubList.append('')
            testResultPassOrNot.append(0)
            break
        elif name == testName[i]:
            countNum = countNum + 1
        else:
            functionName.append(name)
            functionNum.append(countNum)
            stubList.append('')
            testResultPassOrNot.append(0)
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
    count=0
    allcount=0
    stubNum = 0
    for i in range(len(functionName)):
        allcount = allcount + 1
        for j in file_list:
            num = 0
            name = j.split('_test')
            if name[0] == functionName[i]:
                while name[1] != str(num)+'.xls':
                    num = num + 1
                    if num > 50:
                        break
                if num==0:
                    set_color(12)
                    print(str(len(functionName)) + "개 중 " + str(allcount))
                    print("열기:",end='')
                    print(str('testReport/' + projectName + '/Test_Result/'+ name[0] + '_test.xls'))
                    set_color(6)
                book = xw.Book('testReport/' + projectName + '/Test_Result/' + name[0] + '_test'+str(num)+'.xls')
                sheet = xw.sheets['Report']
                testCaseNum.append(int(sheet.range('A11').value))
                text = sheet.range('A1').value
                text = text.strip("테스트 보고서 'SWUTS-F.").strip("'").split("_")
                swddsCode.append(text[0])
                n = 0
                for m in range(1,1000):
                    text1 = sheet.range('A'+str(m)).value
                    if text1 == '번호':
                        num1 = sheet.range('A'+ str(m+1)).value
                        explaintext = sheet.range('C'+ str(m+1)).value
                        caseExplain[count] = str(caseExplain[count] + num1 + '. ' + explaintext + '\n')
                        if n == 0:
                            n = m
                strline = ''
                for k in range(len(caseExplain[count]) - 1):
                    strline = strline + caseExplain[count][k]
                caseExplain[count] = strline
                            #print("입력"+str(n))
                C = sheet.range('A' + str(n + 3)).value
                active = 0
                while active==0:
                    C = sheet.range('A' + str(n + 3)).value
                    B = sheet.range('B' + str(n + 3)).value
                    valueList.append(B)
                    #print(str(n + 3) + "/" + str(C))
                    n = n + 1
                    C = sheet.range('A' + str(n + 4)).value
                    if C == "번호" or C == '4. 테스트 세부 정보':
                        active = 1
                set_color(6)
                text = sheet.range('C15').value
                explain = re.split('[() ]',text)
                if num == 0 and modeAnswer == 'n':
                    print(str('[ '+ functionName[i] + ' ]'), end='')
                    print(" 안에 Stub 함수가 있습니까? ")
                    print("stub이 여러개일 경우, 띄어쓰기로 적어주세요. ex) 1 2 3")
                    print("Num\t/\t Value name")
                    print(" Enter\t/\t<없음>")
                    for p in range(1,len(valueList)+1):
                        print(" "+str(p) + '\t/\t' + str(valueList[p-1]))
                    selectNum = input("숫자 입력 : ").split(" ")
                    if selectNum[0] == '0' or selectNum[0] == '':
                        stubList.append("")
                    else:
                        stubList.append(str(valueList[int(selectNum[0])-1]+', '))
                        for p in range(1,len(selectNum)):
                            stubList[len(stubList)-1] = stubList[len(stubList)-1] + str(valueList[int(selectNum[p])])+', '
                        strline = ''
                        for k in range(len(stubList[i]) - 2):  ## remove ,
                            strline = strline + stubList[i][k]
                        stubList[i] = strline
                set_color(10)
                print("file:",num, end=' ')
                if modeAnswer == 'n':
                    print(functionName[i] + "_" + str(num) + " 는 어떤 타입의 테스트입니까? ")
                    print(" 1\t/\tAnalysis Of Boundary(경계값)")
                    print(" 2\t/\tEquivalence Testing(동등분할)")
                    print(" 3\t/\tFault Injection Testing(결함주입)")
                    print(" Enter\t/\tDevelopment Of Positive(STATEMENT)")
                    answer = str(input("입력 : "))
                else:
                    if "create" in explain or "Create" in explain:
                        answer = ''
                        print("word find.",end=':')
                    elif "boundary" in explain or "Boundary" in explain:
                        answer = '1'
                        print("word find.",end=':')
                    elif "equivalence" in explain or "Equivalence" in explain:
                        answer = '2'
                        print("word find.",end=':')
                    elif "exception" in explain or "Exception" in explain:
                        answer = '3'
                        print("word find.",end=':')
                    else:
                        answer = str(num+1)
                        print("just file number")
                if answer == '':
                    print("Statement")
                    testType.append("4")
                else:
                    if answer == '1':
                        print("Boundary")
                    elif answer == '2':
                        print("Equivalence")
                    elif answer == '3':
                        print("Exception")
                    testType.append(answer)
                set_color(15)
                valueList = []  # resets
                book.app.quit()
                time.sleep(0.4)
                count = count + 1
                set_color(6)
        print()
    xlbook = xw.Book(str("Report.xlsx"))
    sheet = xlbook.sheets['StubList']
    Tester = sheet.range('G3').value
    date = sheet.range('G6').value
    if date == 'now':
        date = datetime.today().strftime('%Y-%m-%d')
    if modeAnswer == 'y':
        print("find Stub")
        numStub = 3
        text = sheet.range('B' + str(numStub)).value
        while text != '':
            print(text)
            numCount = 0
            for i in functionName:
                if i == text:
                    stubList[numCount] = sheet.range('C' + str(numStub)).value
                    break
                else:
                    numCount = numCount + 1
            numStub = numStub + 1
            if numStub > len(functionName):
                break
            text = sheet.range('B' + str(numStub)).value
        print("find NOT")
        numNOT = 3
        text = sheet.range('E' + str(numNOT)).value
        while text != '':
            print(text)
            numCount = 0
            for i in functionName:
                if i == text:
                    testResultPassOrNot[numCount] = 1
                else:
                    numCount = numCount + 1
            numNOT = numNOT + 1
            if numNOT > len(functionName):
                break
            text = sheet.range('E' + str(numNOT)).value
    xlbook.app.quit()
    print("데이터 수집완료")
    print('SWDDS code:\t',end='')
    print(swddsCode)
    print(testCaseNum)
    for a in caseExplain:
        print(a)
    for a in stubList:
        print(a)

    #타입 리스트 변경
    for i in range(len(testType)):
        if testType[i] == '1':    #BND
            testType[i] = DescriptionMsBND
        elif testType[i] == '2':  #EQV
            testType[i] = DescriptionMsEQV
        elif testType[i] == '3':  #FIT
            testType[i] = DescriptionMsFIT
        elif testType[i] == '4':  #state
            testType[i] = DescriptionMsSTA
        else:
            testType[i] = "none"

    #이름에 대한 커버리지 찾기

    print("테스트 보고서 유무확인")
    file_list = os.listdir(str('testReport/'+projectName+'/'))
    if len(file_list) > 0:
        for i in range(len(file_list)):
            if file_list[i] == TCresultName:
                fileOnOff = 1
                resultBook = xw.Book(str('testReport/'+projectName+'/'+TCresultName))
                sheet = resultBook.sheets['Sheet0']
                i = 1
                data = str(sheet[str('A1')].value)
                print("탐색중.....",end='')
                while '함수별 커버리지' != data:
                    i = i + 1
                    data = str(sheet['A' + str(i)].value)
                    if i > 5000:
                        print("not find")
                        break
                    elif i%50==0:
                        print('.',end='')
                print('find')
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
        print('커버리지 보고서 파일없음')
        input()
    #엑셀자동생성
    file_list = os.listdir()
    a = 0
    for i in file_list:
        if i == xl_file:
            a = 1
    if a == 0:
        print("보고서 파일이 없습니다.")
    file_list = os.listdir('Result/')
    for i in file_list:
        os.remove('Result/'+i)
    print("엑셀 생성")
    shutil.copy(str(xl_file), str('Result/' + "보고서.xlsx")) # 결과 출력을 위한 빈파일 생성
    #os.rename(str('Result/' + xl_file), str('Result/' + xl_file))
    xlbook = xw.Book(str('Result/' + "보고서.xlsx"))
    print('출력',end='')
    sheet = xlbook.sheets['Unit_TC']
    Ycell = 0
    xlStartNum = 13
    print(SCNPercent)
    print(SCNTest)
    print(SCNTotal)
    print(BCNPercent)
    print(BCNTest)
    print(BCNTotal)
    p = 0
    for i in range(len(functionName)):
        print("보고서 작성")
        for j in range(int(functionNum[i])):
            sheet.range('B' + str(xlStartNum + Ycell)).value = str('SWUTS-F.' + swddsCode[p] + '_') + str(
                j + 1)  # SWUTS출력
            sheet.range('C' + str(xlStartNum + Ycell)).value = str('SWDDS.' + swddsCode[p])     # SWDDS출력
            sheet.range('D' + str(xlStartNum + Ycell)).value = functionFileName[i]  # 파일이름
            sheet.range('E' + str(xlStartNum + Ycell)).value = functionName[i]                               # unit 이름출력
            sheet.range('F' + str(xlStartNum + Ycell)).value = Tester                                        # 테스터 출력
            sheet.range('G' + str(xlStartNum + Ycell)).value = str("TestCase ID] SWUTS-F." + swddsCode[p] + '_' + str(j + 1) + "\nGoal : " + testType[Ycell])
            sheet.range('L' + str(xlStartNum + Ycell)).value = caseExplain[Ycell]  # 테스트케이스 설명
            sheet.range('Z' + str(xlStartNum + Ycell)).value = date                                          # 날짜
            sheet.range('Q' + str(xlStartNum + Ycell)).value = testCaseNum[Ycell]                            # 테스트케이스 갯수
            if testType[Ycell] == DescriptionMsBND:
                sheet.range('K' + str(xlStartNum + Ycell)).value = "Analysis of boundary values"
            elif testType[Ycell] == DescriptionMsEQV:
                sheet.range('K' + str(xlStartNum + Ycell)).value = "Equivalence testing"
            elif testType[Ycell] == DescriptionMsFIT:
                sheet.range('K' + str(xlStartNum + Ycell)).value = "Error guessing"
                sheet.range('J' + str(xlStartNum + Ycell)).value = "Fault Injection Test"
            elif testType[Ycell] == DescriptionMsSTA:
                sheet.range('K' + str(xlStartNum + Ycell)).value = "Development of positive"
            else:
                sheet.range('K' + str(xlStartNum + Ycell)).value = "Equivalence testing"
            if stubList[i] != '':
                sheet.range('H' + str(xlStartNum + Ycell)).value = "There is no compilation error\nCreate the stub function\n(" + stubList[i] +")" # stub 넣기
            if testResultPassOrNot[i] == 0:
                sheet.range('O' + str(xlStartNum + Ycell)).value = "OK"
            else:
                sheet.range('O' + str(xlStartNum + Ycell)).value = "NOK"
            if fileOnOff == 1:
                print("커버리지 작성")
                if SCNPercent[i] == 'N/A':
                    sheet.range('R' + str(xlStartNum + Ycell)).value = str(SCNPercent[i])
                else:
                    sheet.range('R' + str(xlStartNum + Ycell)).value = str(SCNPercent[i]) + '%'                  # 커버리지 state %
                sheet.range('S' + str(xlStartNum + Ycell)).value = SCNTest[i]                                    # 커버리지 num test
                sheet.range('T' + str(xlStartNum + Ycell)).value = SCNTotal[i]                                   # 커버리지 num total
                if BCNPercent[i] == 'N/A':
                    sheet.range('V' + str(xlStartNum + Ycell)).value = str(BCNPercent[i])
                else:
                    sheet.range('V' + str(xlStartNum + Ycell)).value = str(BCNPercent[i]) + '%'                  # 커버리지 brench %
                sheet.range('W' + str(xlStartNum + Ycell)).value = BCNTest[i]                                    # 커버리지 num test
                sheet.range('X' + str(xlStartNum + Ycell)).value = BCNTotal[i]                                   # 커버리지 num total
            Ycell = Ycell + 1
            print('.', end='')
            p = p + 1
    #xlbook.sheets['Unit_TC'].name = fileName # 시트 이름 변경
    print()
    print("생성완료")
    xlbook.save()
    xlbook.app.quit()
    file_list = os.listdir(str('testReport/'+projectName+'/Test_Result/'))
    print("Test Result xl 파일 가져오는 중.")
    set_color(11)
    print('\x1b[92m')
    p=0
    for i in range(len(functionName)):
        for k in range(functionNum[i]):
            for j in file_list:
                if j == str(functionName[i] + '_test' + str(k) + '.xls'):
                    shutil.copy(str('testReport/' + projectName + '/Test_Result/' + functionName[i] + '_test' + str(
                        k) + '.xls'), str('Result/SWUTS-F.' + swddsCode[p] + '_' + str(k + 1) + '.xls'))
                    print('copy: ' + functionName[i] + '_test' + str(k) + '.xls   \t\t-->\t SWUTS-F.' + swddsCode[
                        p] + '_' + str(k + 1) + '.xls')
                    p =p+1

    set_color(15)
    print('\x1b[97m')
    print("완료")

input('작업끝.  엔터를 눌러주세요.')


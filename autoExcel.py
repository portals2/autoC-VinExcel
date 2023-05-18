'''
해야할 것
1. 드래그는 자동이니까 Ctrl+C&V가 작동
2. 해당 버튼을 누르면 cv작동
3. 각각의 입력을 받은 것을 보여줌
4. 해당 엑셀에 자동 입력

엑셀을 읽어서 a열 행(마직막 번호+1)에 저장시키기

필수셀에 사람이 없으면 밀리는데 이거 해결 해야함
(행 수를 이름으로 고정?)
노트북에서 돌아가야 하니까 
엑셀 주소도 따로 등록

'''

import time
import pyautogui #키보드 C V해주는
import pyperclip #한글 지원
import keyboard #키보드 이벤트 인식
import openpyxl #엑셀


# text.start()

dir = 'C:\\Users\\LIM\\Desktop\\(엑셀파일)메론양식(23년).xlsx'
excel = openpyxl.load_workbook(dir , data_only=True )
excel_ws = excel.active

while True:
    #받는 분
    k = keyboard.read_key()
    n = 0   
    if k == "1":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        a = pyperclip.paste()
        print('1 받는분 :',a)
        for i in range(2, 10000):
            A = excel_ws['A{}'.format(i)]
            if A.value == None:
                excel_ws['A{}'.format(i)] = a
                n = i
                break   
    #받는 분 전화
    if k == "2":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('2 전화번호 :',b)
        excel_ws['B{}'.format(i)] = b
    #주소
    if k == "3":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('3 주소 :',b)
        excel_ws['D{}'.format(i)] = b
    #수량
    if k == "4":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('4 수량 :',b)
        excel_ws['E{}'.format(i)] = b
    #품목명
    if k == "5":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('5 KG :',b)
        excel_ws['F{}'.format(i)] = b
    #메모1
    if k == "6":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('6 메모1 :',b)
        excel_ws['I{}'.format(i)] = b
    #메모2
    if k == "7":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('7 메모2 :',b)
        excel_ws['J{}'.format(i)] = b
    
        
    # 종료 조건문
    if k == "f12":
        break

# 엑셀 저장
excel.save(dir)

# C:\Users\dkfvk\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0
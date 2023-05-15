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


excel = openpyxl.load_workbook('D:\\js\\code\\auto_c&v_in_excel\\출력자료등록_양식_한진-메론양식(L-11) 메모 위치 이동 물어보기.xlsx' , data_only=True )
excel_ws = excel.active

while True:
    #받는 분
    k = keyboard.read_key()
    if k == "1":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        a = pyperclip.paste()
        print('받는분 :',a)
        for i in range(2, 10000):
            A = excel_ws['A{}'.format(i)]
            if A.value == None:
                excel_ws['A{}'.format(i)] = a
                break     
    #받는 분 전화
    if k == "2":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('전화번호 :',b)
        for i in range(2, 10000):
            B = excel_ws['B{}'.format(i)]
            if B.value == None:
                excel_ws['B{}'.format(i)] = b
                break
    #주소
    if k == "3":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('주소 :',b)
        for i in range(2, 10000):
            B = excel_ws['D{}'.format(i)]
            if B.value == None:
                excel_ws['D{}'.format(i)] = b
                break
    #수량
    if k == "4":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('수량 :',b)
        for i in range(2, 10000):
            B = excel_ws['E{}'.format(i)]
            if B.value == None:
                excel_ws['E{}'.format(i)] = b
                break
    #품목명
    if k == "5":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('KG :',b)
        for i in range(2, 10000):
            B = excel_ws['F{}'.format(i)]
            if B.value == None:
                excel_ws['F{}'.format(i)] = b
                break
    #메모1
    if k == "6":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('메모1 :',b)
        for i in range(2, 10000):
            B = excel_ws['I{}'.format(i)]
            if B.value == None:
                excel_ws['I{}'.format(i)] = b
                break
    #메모2
    if k == "7":
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        b = pyperclip.paste()
        print('메모2 :',b)
        for i in range(2, 10000):
            B = excel_ws['J{}'.format(i)]
            if B.value == None:
                excel_ws['J{}'.format(i)] = b
                break
    
        
    # 종료 조건문
    if k == "f12":
        break

# 엑셀 저장
excel.save('D:\\js\\code\\auto_c&v_in_excel\\출력자료등록_양식_한진-메론양식(L-11) 메모 위치 이동 물어보기.xlsx')

# C:\Users\dkfvk\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0
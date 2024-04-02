import tkinter as tk
from tkinter import messagebox as mb
from tkinter import *
import time
from cryptography.fernet import Fernet
from datetime import timedelta, datetime
import pyautogui
pyautogui.FAILSAFE = False
import win32com.client as win32c 
import os, time
import shutil, psutil

import re
#----------------------------------------------------------------
# Description:
#  - 엑셀컬럼(Alphabet주소)을 column index(start=1)로 변경하는 함수
#----------------------------------------------------------------
# 인덱스 방식 ('A': 0, 'B': 1, ...)
def ColToNum0(colStr):
    """ Convert base26 column string to number. """
    expn = 0
    colNum = 0
    for char in reversed(colStr):
        colNum += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    #return colNum
    return colNum-1

# 엑셀방식 ('A': 1, 'B': 2, ...)

def ColToNum1(colStr):
    """ Convert base26 column string to number. """
    expn = 0
    colNum = 0
    for char in reversed(colStr):
        colNum += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    #return colNum-1
    return colNum
#----------------------------------------------------------------
# Description:
#  - 엑셀컬럼column index(start=1)을 Alphabet주소로 변경하는 함수
#----------------------------------------------------------------
# 인데스방식 (0: 'A', 1: 'B', ...)
def Num0ToCol(n):
    string = ""
    n = n+1
    while n > 0 : 
        n, remainder = divmod(n, 26)
        if(remainder == 0) :
            n = n-1
            string = 'Z' + string   
        else :
            string = chr(ord('A') + remainder -1) + string
    return string

# 엑셀 방식 (1: 'A', 2: 'B', ...)
def Num1ToCol(n):
    string = ""
    #n = n+1
    while n > 0 : 
        n, remainder = divmod(n, 26)
        if(remainder == 0) :
            n = n-1
            string = 'Z' + string   
        else :
            string = chr(ord('A') + remainder -1) + string
    return string

def GetRangeAsList(ws, rng, o_type="value"):
    r1, r2 = re.findall("[0-9]+", rng)
    c1, c2 = re.findall("[A-Z]+", rng)
    rows = list(range(int(r1), int(r2)+1))
    cols = list(range(ColToNum1(c1), ColToNum1(c2)+1))
    
    lines = []
    for row in rows:
        line=[]
        for col in cols:
            if o_type == 'value':
                line.append(ws.Cells(row,col).value)
            else:
                line.append(ws.Cells(row,col).text)
        lines.append(line)
    return lines 

star_text1 = ["*"]
star_text2 = ["*"]
def callback(e1, e2):
    if mb.askyesno('비번변경 확인', '신규비밀번호로 변경할까요?'):
        pw1 = e1.get()
        pw2 = e2.get()
        if pw1 == pw2:
            cipherInfo = []
            #Get current time
            curr_time = datetime.now().strftime('%Y-%m-%d %H:%M')
            #print('curr_time: ', curr_time)
            cipherInfo.append(curr_time)
            #Generate enc
            key = b'insert key'            
            cipher = Fernet(key)
            enc_byteobj = cipher.encrypt(pw1.encode())
            cipher_text = enc_byteobj.decode('utf-8')
            cipherInfo.append(cipher_text)
            #Update excel
            excel = win32c.Dispatch("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False
            excel.WindowState = -4137 # set this number meaning full size
            configfile = r'기준정보 파일 경로 입력' 
            wb = excel.Workbooks.Open(configfile)
            ws = wb.Sheets('Config')
            max_rows = len(ws.UsedRange.Rows)
            rng = 'A1:'+'C'+str(max_rows)
            lines = GetRangeAsList(ws, rng, o_type="text")
            for row, line in enumerate(lines):
                if line[1] == 'M-PWD':
                    ws.Cells(row+1,3).value = cipher_text
                    ws.Cells(row+2,3).value = curr_time
                    break
            wb.Save()
            excel.Workbooks.Close()
            excel.Quit()
            time.sleep(1)
            mb.showwarning('변경상태 알림', '신규비빌번호로 변경되었습니다.\n\n'+ '확인은\n  '+configfile + '\n파일을 참조해 주세요.')
        else:
            mb.showwarning('오류 알림', '비밀번호가 일치하지 않습니다.')
        time.sleep(1)
        
        
    else:
        mb.showwarning('변경상태 알림', '변경이 취소되었습니다.')
        
def showstar1(e1, star_text1):
    #print('star_text1: ', star_text1)
    if star_text1[0] == "*":
      e1.configure(show="")
      star_text1[0] = ""
    else:
      e1.configure(show="*")
      star_text1[0] = "*"

def showstar2(e2, star_text2):
    #print('star_text2: ', star_text2)
    if star_text2[0] == "*":
      e2.configure(show="")
      star_text2[0] = ""
    else:
      e2.configure(show="*")
      star_text2[0] = "*"
   
#------------------------------------------------------------------------------------------------
# Driver Code
#------------------------------------------------------------------------------------------------
#if __name__ == "__main__":
window = Tk()
window.geometry('800x500')
#window.eval('tk::PlaceWindow . center')
#window.attributes("-alpha", 1.0) #transparent
#offset_y = int(window.geometry().rsplit('+', 1)[-1])
#bar_height = window.winfo_rooty() - offset_y+5
window.iconbitmap(r'icon_path')
window.title('계정비번 변경툴')
l00 = Label(window, text="신규비번 입력 >>>", width=15, font=("맑은고딕", 12), pady=30)
g01 = l00.grid(row=0)
l10 = Label(window, text="한번 더 입력 >>>", width=15, font=("맑은고딕", 12))
g11 = l10.grid(row=1)

#https://cs111.wellesley.edu/archive/cs111_fall14/public_html/labs/lab12/tkintercolor.html 참조
e1 = Entry(window, text = '', font=("맑은고딕", 16), show="*", width=30, bg="bisque", fg='green4')
e2 = Entry(window, text = '', font=("맑은고딕", 16), show="*", width=30, bg="bisque", fg='green4')

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)

b01 = Button(window, text='*', command=lambda:showstar1(e1, star_text1), width=1, height=1, font=("맑은고딕", 11), fg='red')
b01.grid(row=0, column=2, sticky=tk.W)
b11 = Button(window, text='*', command=lambda:showstar2(e2, star_text2), width=1, height=1, font=("맑은고딕", 11), fg='red')
b11.grid(row=1, column=2, sticky=tk.W)
b31 = Button(window, text='변경/등록', command=lambda:callback(e1, e2), width=17, height=1, font=("맑은고딕", 11))
b31.grid(row=3, column=0, sticky=tk.W, padx = 40, pady=40)
b32 = Button(window, text='종료', command=window.quit, width=17, height=1, font=("맑은고딕", 11))
b32.grid(row=3, column=1, sticky=tk.W, padx = 0, pady=40)  

window.mainloop()

import pyautogui as pag
import time
import webbrowser
import win32com.client as win32
import os
import pyautogui

time.sleep(5)
webbrowser.open('http://ors/ors/atm/promise.html', new=2, autoraise=True)
time.sleep(30)
pag.click(1506, 363, 1, 2, 'left')
time.sleep(15)
pag.click(1508, 299, 1, 2, 'left')
time.sleep(25)
pag.click(1535, 137, 1, 2, 'left')
time.sleep(10)
pag.hotkey('alt', 'f8')
time.sleep(5)
pag.hotkey('enter')
time.sleep(45)
pag.hotkey('f12')
time.sleep(1)
pag.typewrite('ORS.xlsx', interval=0.25)
time.sleep(10)
pag.hotkey('enter')
time.sleep(5)
pyautogui. screenshot(r'C:\Users\u_180u6\Downloads\ORS.jpg',region=(0,0, 1800, 1000))
time.sleep(5)
pag.hotkey('alt', 'f4')
time.sleep(5)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = '' # Почта
mail.Subject = 'Расчет ORS'
mail.Body = 'Расчет ORS на текущую Дату'
mail.HTMLBody = '<h2>Расчет ORS на текущую Дату</h2>'

attachment  = (r'C:\Users\u_180u6\Downloads\ORS.xlsx')
attachment1  = (r'C:\Users\u_180u6\Downloads\ORS.jpg')
mail.Attachments.Add(attachment)
mail.Attachments.Add(attachment1)

mail.Send()

time.sleep(30)

path = (r'C:\Users\u_180u6\Downloads\ORS.xlsx')
path1 = (r'C:\Users\u_180u6\Downloads\ORS.jpg')
os.remove(path)
os.remove(path1)
from selenium import webdriver
from selenium.webdriver.common.by import By
import win32com.client as win32
import py_win_keyboard_layout
import pyautogui as pag
import pyautogui
import logging
import time
import os

logging.basicConfig(filename = "log.log" , level=logging.INFO , format = '%(asctime)s %(levelname)s %(funcName)s || %(message)s')

def Ors():  # Работа с ORS
    link = 'http://ors/ors/atm/promise.html'
    driver = webdriver.Edge()
    driver.maximize_window()
    time.sleep(5)
    driver.get(link)
    time.sleep(10)
    driver.implicitly_wait(240)
    button = driver.find_element(By.ID, 'searchButton').click()  # Обновить
    driver.implicitly_wait(360)
    button1 = driver.find_element(By.ID, 'exportButtonDetail').click()  # Детальный отчет
    time.sleep(240)
    pag.hotkey('ctrl', 'j')  # Загрузки
    time.sleep(3)
    pag.hotkey('enter')
    time.sleep(5)

def Excel():  # Работа с excel
    py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)  # Смена языка на EN
    time.sleep(2)
    pag.hotkey('alt', 'f8')  # Макрос ORS
    time.sleep(2)
    pag.hotkey('enter')
    time.sleep(60)
    pag.hotkey('f12')  # Сохранение
    time.sleep(1)
    pag.typewrite('ORS.xlsx', interval=0.1)  # Переименование
    time.sleep(5)
    pag.hotkey('enter')
    time.sleep(2)
    pyautogui. screenshot(r'C:\Users\u_180u6\Downloads\ORS.jpg', region=(0, 0, 1800, 1000))  # Скриншот
    time.sleep(5)
    os.system("taskkill /f /im EXCEL.exe")  # Закрытие
    time.sleep(5)

def Outlook():  # Отправка в Outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = '' # Почта
    mail.Subject = 'Расчет ORS'
    mail.Body = 'Расчет ORS на текущую Дату'
    mail.HTMLBody = '<h2>Расчет ORS на текущую Дату</h2>'
    attachment = (r'C:\Users\u_180u6\Downloads\ORS.xlsx')
    attachment1 = (r'C:\Users\u_180u6\Downloads\ORS.jpg')
    mail.Attachments.Add(attachment)
    mail.Attachments.Add(attachment1)
    mail.Send() # Отправка почты

def Delete():  # Удаление лишнего
    time.sleep(10)
    path = (r'C:\Users\u_180u6\Downloads\ORS.xlsx')
    path1 = (r'C:\Users\u_180u6\Downloads\ORS.jpg')
    os.remove(path)
    os.remove(path1)

start = True
while start:
    try:
        Ors()
        Excel()
        start = False
        Outlook()
        Delete()
    except:
        logging.exception(Ors)
        os.system("taskkill /f /im msedgedriver.exe")
        os.system("taskkill /f /im EXCEL.exe") 
        time.sleep(10)
        pass

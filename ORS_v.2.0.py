from selenium.webdriver.common.by import By
from selenium import webdriver
from pathlib import *
import os , glob , time , logging , pyautogui , win32com.client , py_win_keyboard_layout , win32com.client as win32

logging.basicConfig(filename = "log.log" , level=logging.INFO , format = '%(asctime)s %(levelname)s %(funcName)s || %(message)s')
py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)  # Смена языка на EN

def ORS():  # Работа с ORS
    start_ors = True
    while start_ors:
        try:
            link = ('http://ors/ors/atm/promise.html')
            driver = webdriver.Edge()
            driver.maximize_window()
            time.sleep(5)
            driver.get(link)
            time.sleep(10)
            driver.implicitly_wait(240)
            driver.find_element(By.ID, 'searchButton').click()  # Обновить
            driver.implicitly_wait(360)
            driver.find_element(By.ID, 'exportButtonDetail').click()  # Детальный отчет
            start_ors = False
            time.sleep(240)
            logging.info('-----OK-----')
        except:
            logging.exception(ORS)
            os.system("taskkill /f /im msedgedriver.exe")
            os.system("taskkill /f /im msedge.exe")
            time.sleep(15)

def EXL(): # Работа с EXl
    try:
        fileors = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/', 'detail_1672*.xlsx')) # Поиск
        for ors in fileors:
            pass
            xlApp = win32com.client.DispatchEx('Excel.Application')
            wb = xlApp.Workbooks.Open(ors)
            xlApp.Visible = True
            xlApp.Run('PERSONAL.XLSB!ORS_v_4_1') # Макрос
            time.sleep(60)
            pyautogui.screenshot(r'C:\Users\u_180u6\Downloads\ORS.jpg', region=(0, 0, 1800, 1000))  # Снимок    
            wb.Save() # Сохранение
            xlApp.Quit() # Выход
            logging.info('-----OK-----')
            time.sleep(10)
    except:
            logging.exception(EXL)
            os.system("taskkill /f /im EXCEL.exe")
            xlApp.Quit()
            
def Out():  # Отправка в Outlook
    try:
        fileout = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/ORS_*.xlsx')) # Поиск
        for out in fileout:
            pass
        filejpg = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/ORS*.jpg')) # Поиск
        for jpg in filejpg:
            pass
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = '' # Почта
            mail.Subject = 'Расчет ORS'
            mail.Body = 'Расчет ORS на текущую Дату'
            mail.HTMLBody = '<h2>Расчет ORS на текущую Дату</h2>'
            mail.Attachments.Add(out)
            mail.Attachments.Add(jpg)
            mail.Send() # Отправка почты
            logging.info('-----OK-----')
    except:
        logging.exception(Out)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = '' # Почта
        mail.Subject = 'Неудача подчета ORS'
        mail.Body = 'Неудача подчета ORS'
        mail.HTMLBody = '<h2>Неудача подчета ORS</h2>'
        mail.Send() # Отправка почты    

def Delete():  # Удаление лишнего
    try:
        time.sleep(300)
        fileout = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/ORS_*.xlsx')) # Поиск
        for out in fileout:
            pass
        filejpg = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/ORS*.jpg')) # Поиск
        for jpg in filejpg:
            pass
        fileors = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/', 'detail_1672*.xlsx')) # Поиск
        for ors in fileors:
            pass
        os.remove(ors)
        os.remove(out)
        os.remove(jpg)
        logging.info('-----OK-----')
    except:
        logging.exception(Delete)

start = (ORS(), EXL(), Out(), Delete()) # Поехали ;)
from selenium.webdriver.common.by import By
from selenium import webdriver
from pathlib import *
import os , glob , time , logging , py_win_keyboard_layout , win32com.client as win32

logging.basicConfig(filename = "log.log" , level=logging.INFO , format = '%(asctime)s %(levelname)s %(funcName)s || %(message)s') # Логи
py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)  # Смена языка на EN

def ORS():  # Работа с ORS
    start_ors = True
    while start_ors:
        try:
            driver = webdriver.Edge()
            driver.maximize_window()
            time.sleep(3)
            driver.get('http://ors/ors/atm/promise.html')
            time.sleep(5)
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
            xlApp = win32.Dispatch('Excel.Application')
            wb = xlApp.Workbooks.Open(ors)
            xlApp.Visible = False
            xlApp.Run('PERSONAL.XLSB!ORS_v_4_1') # Макрос
            time.sleep(45)  
            wb.Save() # Сохранение
            wb.Worksheets("Total").ExportAsFixedFormat(0, 'C:/Users/u_180u6/Downloads/ORS.pdf') # Сохранение в PDF
            xlApp.Quit() # Выход
            logging.info('-----OK-----')
            time.sleep(5)
    except:
            logging.exception(EXL)
            os.system("taskkill /f /im EXCEL.exe")
            
def Out():  # Отправка в Outlook
    try:
        fileout = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/ORS_*.xlsx')) # Поиск
        for out in fileout:
            pass
        fileors = glob.glob(os.path.join('C:/Users/u_180u6/Downloads', 'ORS.pdf')) # Поиск
        for pdf in fileors:
            pass
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = '' # Почта
            mail.Subject = 'Расчет ORS'
            mail.Body = 'Расчет ORS на текущую Дату'
            mail.HTMLBody = '<h2>Расчет ORS на текущую Дату</h2>'
            mail.Attachments.Add(out)
            mail.Attachments.Add(pdf)
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
        time.sleep(60)
        fileout = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/ORS_*.xlsx')) # Поиск
        for out in fileout:
            pass
        fileors = glob.glob(os.path.join('C:/Users/u_180u6/Downloads', 'ORS.pdf')) # Поиск
        for pdf in fileors:
            pass
        fileors = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/', 'detail_1672*.xlsx')) # Поиск
        for ors in fileors:
            pass
        os.remove(ors)
        os.remove(out)
        os.remove(pdf)
        logging.info('-----OK-----')
    except:
        logging.exception(Delete)

start = (ORS(), EXL(), Out(), Delete()) # Поехали ;)
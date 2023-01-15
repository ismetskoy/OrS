from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
from selenium import webdriver
from threading import Thread
from pathlib import *
import os , glob , time , logging  , win32com.client as win32

logging.basicConfig(filename = "log.log" , level=logging.INFO , format = '%(asctime)s %(levelname)s %(funcName)s || %(message)s') # Логи

day =  (datetime.now().replace(microsecond=0) + timedelta(days=-1))  # Дата

def TimeEXL(): # Kill EXCEL
    time.sleep(120)
    os.system("taskkill /f /im EXCEL.exe")

def Poisk(): # Ожидание загрузки файлы detail
    file_time=time.time()
    while (time.time() - file_time) < 300: 
        filenames = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/', 'detail_*.xlsx'))
        if len(filenames) > 0 :
            time.sleep(3)
            break
    else:
        logging.exception(ORS)
        os.system("taskkill /f /im msedgedriver.exe")
        os.system("taskkill /f /im msedge.exe")
        time.sleep(15)
        ORS()

def ORS():  # Работа с сайтом ORS
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
            time.sleep(5)
            driver.implicitly_wait(360)
            driver.find_element(By.ID, 'exportButtonDetail').click()  # Детальный отчет
            start_ors = False
            Poisk()
            logging.info('-----OK-----')
        except:
            logging.exception(ORS)
            os.system("taskkill /f /im msedgedriver.exe")
            os.system("taskkill /f /im msedge.exe")
            time.sleep(15)

def EXL(): # Работа с EXl
    try:
        filedet = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/', 'detail_*.xlsx')) # Поиск
        for det in filedet:
            pass
        xlApp = win32.Dispatch('Excel.Application')
        wb = xlApp.Workbooks.Open(det)
        xlApp.Visible = False
        xlApp.Run('PERSONAL.XLSB!ORS_v_4_1') # Макрос
        time.sleep(60)  
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
        fileors = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/', 'ORS*.xlsx'))
        for ors in fileors:
            pass
        filepdf = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/', 'ORS*.pdf'))
        for pdf in filepdf:
            pass
        filejpg = glob.glob(os.path.join('C:/Users/u_180u6/Downloads/', 'ORS*.jpg'))
        for jpg in filejpg:
            pass
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = ''  # Почта
        mail.Subject = 'Расчет ORS' 
        mail.Body = 'Расчет ORS на Дату: {day}'
        mail.HTMLBody =  "<html><body><h3>Расчет ORS на Дату: {day} <br></h3></body></html>".format(day=day)
        mail.Attachments.Add(ors)
        mail.Attachments.Add(pdf)
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
        time.sleep(30)
        filedel = glob.glob(os.path.join
            ('C:/Users/u_180u6/Downloads/', 'ORS*.xlsx')) + glob.glob(os.path.join
            ('C:/Users/u_180u6/Downloads/', 'ORS*.jpg')) + glob.glob(os.path.join
            ('C:/Users/u_180u6/Downloads/', 'ORS*.pdf')) + glob.glob(os.path.join
            ('C:/Users/u_180u6/Downloads/', 'detail_*.xlsx')) 
        for delete in filedel:
            os.remove(delete)
            pass
        logging.info('-----OK-----')
    except:
        logging.exception(Delete)

start = (ORS(), Thread(target=TimeEXL).start(), EXL(), Out(), Delete()) # Поехали ;)
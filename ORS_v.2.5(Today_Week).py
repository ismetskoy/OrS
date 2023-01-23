from selenium.webdriver.common.by import By
from datetime import datetime, timedelta, date
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from threading import Thread
from pathlib import *
import os , glob , time , logging  , win32com.client as win32

logging.basicConfig(filename = "log.log" , level=logging.INFO , format = '%(asctime)s %(levelname)s %(funcName)s || %(message)s') # Логи

week = f"{date.today() - timedelta(date.today().weekday()):%d.%m.%Y}" " --- " f"{datetime.now() + timedelta(days=-1):%d.%m.%Y}"" 🚀 " 

start_today = f'{date.today() - timedelta(date.today().weekday()):%d.%m.%Y}'

def TimeEXL(): # Kill EXCEL
    time.sleep(120)
    os.system("taskkill /f /im EXCEL.exe")

def Poisk(): # Ожидание загрузки файлы detail
    file_time=time.time()
    while (time.time() - file_time) < 300:
        filenames = glob.glob(os.path.join('C:/Users/*/Downloads/', 'detail_*.xlsx'))
        if len(filenames) > 0 :
            time.sleep(3)
            break
    else:
        logging.exception(ORS)
        os.system("taskkill /f /im msedgedriver.exe")
        time.sleep(15)
        ORS()

def ORS():  # Работа с сайтом ORSe
    start_ors = True
    while start_ors:
        try:
            driver = webdriver.Edge()
            driver.maximize_window()
            time.sleep(3)
            driver.get('http://ors/ors/atm/promise.html')
            time.sleep(5)
            driver.implicitly_wait(240)
            driver.find_element(By.ID, 'dateFrom').click()
            time.sleep(5)
            driver.find_element(By.ID, 'dateFrom').send_keys(start_today)
            time.sleep(5)
            driver.find_element(By.ID, 'dateFrom').send_keys(Keys.RETURN)
            driver.find_element(By.ID, 'dateFrom').send_keys(Keys.RETURN)
            start_ors = False
            time.sleep(5)
            Poisk()
            logging.info('-----OK-----')
        except:
            logging.exception(ORS)
            os.system("taskkill /f /im msedgedriver.exe")
            time.sleep(15)

def EXL(): # Работа с EXl
    try:
        filedet = glob.glob(os.path.join('C:/Users/*/Downloads/', 'detail_*.xlsx')) # Поиск
        for det in filedet:
            pass
        xlApp = win32.Dispatch('Excel.Application')
        wb = xlApp.Workbooks.Open(det)
        xlApp.Visible = False
        xlApp.Run('PERSONAL.XLSB!ORS_v_4_2') # Макрос
        time.sleep(60)  
        wb.Save() # Сохранение
        xlApp.Quit() # Выход
        logging.info('-----OK-----')
        time.sleep(5)
    except:
        logging.exception(EXL)
        os.system("taskkill /f /im EXCEL.exe")
        time.sleep(10)
        EXL()
            
def Out():  # Отправка в Outlook
    try:
        fileors = glob.glob(os.path.join('C:/Users/*/Downloads/', 'ORS*.xlsx'))
        for ors in fileors:
            pass
        filejpg = glob.glob(os.path.join('C:/Users/*/Downloads/', 'ORS*.jpg'))
        for jpg in filejpg:
            pass
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = '' # Отправка почты
        mail.Subject = 'Расчет ORS в динамике' 
        mail.Body = 'Расчет ORS на Дату: {week}'
        mail.HTMLBody =  "<html><body><h2>Расчет ORS на Дату: {week} <br></h2><img src=""cid:MyId1""></body></html>".format(week=week)
        attachment = mail.Attachments.Add(jpg)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
        mail.Attachments.Add(ors)
        mail.Send() # Отправка почты
        logging.info('-----OK-----')
    except:
        logging.exception(Out)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = '' # Отправка почты
        mail.Subject = 'Неудача подчета ORS'
        mail.Body = 'Неудача подчета ORS за {week}'
        mail.HTMLBody = "<html><body><h2>Неудача подчета ORS за {week}<br></h2></body></html>".format(week=week)
        mail.Send() # Отправка почты    

def Delete():  # Удаление лишнего
    try:
        time.sleep(10)
        filedel = glob.glob(os.path.join
            ('C:/Users/*/Downloads/', 'ORS*.xlsx')) + glob.glob(os.path.join
            ('C:/Users/*/Downloads/', 'ORS*.jpg'))+ glob.glob(os.path.join
            ('C:/Users/*/Downloads/', 'detail_*.xlsx')) 
        for delete in filedel:
            os.remove(delete)
            pass
        logging.info('-----OK-----')
    except:
        logging.exception(Delete)

start = (Delete(), ORS(), Thread(target=TimeEXL).start(), EXL(), Out(), Delete()) # Поехали ;)

from selenium.webdriver.common.by import By
from datetime import datetime, timedelta, date
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from threading import Thread
from pathlib import *
import os , glob , time , logging  , win32com.client as win32 , subprocess

log_dir = 'C:\\ORS\\log' # Проверка Пути для логов

if not os.path.exists(log_dir):
    os.makedirs(log_dir, exist_ok=True)

log_file = os.path.join(log_dir, "today_week.log") # Запись лога в файл

logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s %(levelname)s %(funcName)s || %(message)s')  # Конфигурация логов

week = f"{date.today() - timedelta(date.today().weekday()):%d.%m.%Y}" " --- " f"{datetime.now() + timedelta(days=-1):%d.%m.%Y}"" 📈 " 
weeks = f"{datetime.now():%U}"
start_today = f'{date.today() - timedelta(date.today().weekday()):%d.%m.%Y}'

def iexplore(): # Открытие iexplore для макроса
    SW_MINIMIZE = 6
    info = subprocess.STARTUPINFO()
    info.dwFlags = subprocess.STARTF_USESHOWWINDOW
    info.wShowWindow = SW_MINIMIZE
    subprocess.Popen(r'C:\Program Files\Internet Explorer\iexplore.exe', startupinfo=info)

def TimeKill(): # Закрытие программы
    file_time=time.time()
    while (time.time() - file_time) < 555: 
        filenames = glob.glob(os.path.join('C:/Users/*/Downloads/', 'detail_*.xlsx'))
        if len(filenames) < 1 :
            logging.info('Файл detail удален, закрытие программы')
            time.sleep(5)
            os.system("taskkill /f /im msedgedriver.exe")
            os.system("taskkill /f /im day.exe")
            break            
    else:
        logging.error("Превышение времени программы")
        time.sleep(5)
        os.system("taskkill /f /im msedgedriver.exe")
        os.system("taskkill /f /im day.exe")

def Poisk():  # Поиск файла detail
    file_time = time.time()
    while (time.time() - file_time) < 300: 
        try:
            filenames = glob.glob(os.path.join('C:/Users/*/Downloads/', 'detail_*.xlsx'))
        except Exception as e:
            logging.error("Произошла ошибка при поиске файла: %s", e)
            break
        if len(filenames) > 0:
            time.sleep(3)
            logging.info('Файл найден: %s', filenames[0])
            break
    else:
        logging.warning("Время поиска файла истекло")
        os.system("taskkill /f /im msedgedriver.exe")
        time.sleep(5)
        Delete()
        ORS()

def ORS():  # Работа с сайтом ORSe
    start_ors = True
    while start_ors:
        try:
            driver = webdriver.Edge()
            driver.minimize_window()
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
            logging.info('Выгрузка файла прошла')
            Poisk()
            driver.quit()
        except:
            logging.exception(ORS)
            driver.quit()
            os.system("taskkill /f /im msedgedriver.exe")

def EXL(): # Работа с EXl
    try:
        filedet = glob.glob(os.path.join('C:/Users/*/Downloads/', 'detail_*.xlsx')) # Поиск
        xlApp = win32.Dispatch('Excel.Application')
        wb = xlApp.Workbooks.Open(filedet[0])
        xlApp.Visible = False
        xlApp.Run('ORS.xlsb!V_6_2') # Макрос 
        wb.Save() # Сохранение
        xlApp.Quit() # Выход
        logging.info('Макрос выполнен')
        os.system("taskkill /f /im iexplore.exe")
    except:
        logging.exception(EXL)
        os.system("taskkill /f /im iexplore.exe")
        os.system("taskkill /f /im EXCEL.exe")
            
def Out():  # Отправка в Outlook
    try:
        fileors = glob.glob(os.path.join('C:/Users/*/Downloads/', 'ORS*.xlsx'))
        filejpg = glob.glob(os.path.join('C:/Users/*/Downloads/', 'ORS*.jpg'))
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = '' # Отправка почты
        mail.Subject = 'Расчет ORS в динамике'
        mail.Body = 'Расчет ORS на Дату: {week} неделя {weeks}'.format(week=week, weeks=weeks)
        mail.HTMLBody =  "<html><body><h2>Расчет ORS на Дату: {week} неделя {weeks}<br></h2><img src=""cid:MyId1""></body></html>".format(week=week, weeks=weeks)
        attachment = mail.Attachments.Add(filejpg[0])
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
        mail.Attachments.Add(fileors[0])
        mail.Send() 
        logging.info('Отправка почты выполнена')
    except:
        logging.exception(Out)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = '' # Отправка почты
        mail.Subject = 'Неудача подчета ORS'
        mail.Body = 'Неудача подчета ORS за {week}'
        mail.HTMLBody = "<html><body><h2>Неудача подчета ORS за {week}<br></h2></body></html>".format(week=week)
        mail.Send() # Отправка почты    

def Delete():  # Удаление файлов
    try:
        time.sleep(5)
        os.system("taskkill /f /im EXCEL.exe")
        filedel = glob.glob(os.path.join
            ('C:/Users/*/Downloads/', 'ORS*.xlsx')) + glob.glob(os.path.join
            ('C:/Users/*/Downloads/', 'ORS*.jpg')) + glob.glob(os.path.join
            ('C:/Users/*/Downloads/', 'detail_*.xlsx'))
        for delete in filedel:
            os.remove(delete)
            pass
        logging.info('Файлы удаленны')
    except:
        logging.exception(Delete)

start = (Delete(), ORS(), Thread(target=TimeKill).start(), Thread(target=iexplore).start(), EXL(), Out(), Delete()) # Поехали ;)

#!/usr/bin/env python
# coding: utf-8

# # Глобальные переменные

# In[1]:


# Глобальные параметры, стандартные
list_not_data_for_view = []
loading_path_part = r'C:/Users/User/Desktop/py scripts/АИСОИП статусы kpi'
path_load_finish = r'C:/Users/User/Desktop/py scripts/АИСОИП отмены kpi'     # Где хранятся отмены и законченные производства.
path_crm = r"\\192.168.1.251\kpi"
run_prod = True

dict_log_pas = {"kpi":{'log':'901021350973_210540032049', 'pas':'Kk12345%'}}

count_row_in_load_excel = 10_000
cnt_row_load_and_id = {100:0, 1000:1, 10_000:2, 20_000:3, 50_000:-1}
id_load_type = cnt_row_load_and_id[count_row_in_load_excel]


# ## Импорты

# In[2]:


from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from tqdm import tqdm
import time
import pandas as pd
from bs4 import BeautifulSoup
import os
import pickle
import math
from tqdm import tqdm

from datetime import date
from dateutil.relativedelta import relativedelta
from datetime import date
from datetime import datetime, timedelta
import numpy as np
from pathlib import Path

import sys
import traceback
from selenium.webdriver.support.wait import WebDriverWait
from dateutil.relativedelta import relativedelta
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
import math


# In[3]:


def find_cnt_all_row(driver):
    '''
    example
    test_cnt = 1-100 из 4443
    cnt_all_row = 4443
    '''
    
    css_el = '.v-data-iterator.mt-4.mt-md-9 .v-data-footer__pagination'
    test_cnt = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_el))
    ).text
    time.sleep(1)
    
    cnt_all_row = int(test_cnt.split(' из ')[-1])
    cnt_iter = math.ceil(cnt_all_row / count_row_in_load_excel)
    print(test_cnt)
    return cnt_iter, test_cnt

def click_next_excel_50k(driver):
    css_el = '.v-data-iterator.excel-footer .v-data-footer .v-data-footer__icons-after .v-icon.notranslate.material-icons.theme--light'
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_el))
    )
    field.click()
    time.sleep(10)
    
def click_load1(driver):
    css_el = '.d-flex.justify-end .v-btn.v-btn--icon.v-btn--round.theme--light.v-size--default .v-btn__content'
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_el))
    )
    field.click()
    time.sleep(2)
    
def load_need_value(driver):
    """
    Выбираю по сколько дел буду скачивать.
    """
    # Перехожу на нужную вкладку
    css_el = '.v-input.v-input--hide-details.v-input--is-label-active.v-input--is-dirty.theme--light.v-text-field.v-text-field--is-booted.v-select .v-icon.notranslate.material-icons.theme--light'
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_el))
    )
    field.click()
    
    css_el = '.v-list.v-select-list.v-sheet.theme--light.theme--light .v-list-item__content'
    field = driver.find_elements(By.CSS_SELECTOR, css_el)[id_load_type]
    field.click()
    time.sleep(2)
    
def click_load2(driver):
    css_el = '.v-btn.v-btn--is-elevated.v-btn--has-bg.theme--light.v-size--small.primary .v-btn__content'
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_el))
    )
    field.click()
    time.sleep(100)

def date_to_str(date_):
    day_   = str(date_.day)
    month_ = str(date_.month)
    year_ = str(date_.year)
    if len(day_) == 1:
        day_ = '0' + day_
    if len(month_) == 1:
        month_ = '0' + month_
    return f'{day_}.{month_}.{year_}'

def to_site(loading_path_all):
    # Вхожу на сайт
    url = 'https://aisoip.adilet.gov.kz/cabinet/exec-productions'

    # Добавляю папку сохранения 
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    prefs = {"profile.default_content_settings.popups": 0,
             "download.default_directory": (loading_path_all + '/').replace('/', '\\'),
             "directory_upgrade": True,
             "profile.managed_default_content_settings.images": 2}
    options.add_experimental_option("prefs", prefs)
    
    # Подключение к сайту
    driver = webdriver.Chrome(options=options)
    
    driver.implicitly_wait(10)
    driver.get(url)
    time.sleep(2)
    
    return driver
    
def write_login(driver, name_login):
    '''Ввожу логин пароль'''
    login_global = dict_log_pas[name_login]['log']
    pasword_global = dict_log_pas[name_login]['pas']    
    
    login  = login_global
    pasword = pasword_global
    
    time.sleep(0.5)
    
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="usernameUserInput"]'))
    )
    field.send_keys(login)
    
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.ID, "password"))
    )
    field.send_keys(pasword)
    
    time.sleep(0.5)
    
    css_el = '.buttons>.ui.primary.button.fluid'
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_el))
    )
        
    field.click()
    time.sleep(0.5)
    
def to_targer_sheet_1(driver):
    # Перехожу на нужную вкладку
    # исполнительные производства
    css_el = '.cabinet-header>.navigation>:nth-child(2)'
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_el))
    )
    field.click()
       
def to_targer_sheet_2(driver):
    # Перехожу на нужную вкладку
    css_el = '.v-slide-group__content.v-tabs-bar__content>:nth-child(3)'
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_el))
    )
    field.click()
    
def set_100_row(driver):
    time.sleep(0.5)
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '.v-input.v-input--hide-details.v-input--is-label-active.v-input--is-dirty.theme--light.v-text-field.v-text-field--is-booted.v-select .v-select__selection.v-select__selection--comma'))
    )
    field.click()
    
    field = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '.v-list.v-select-list.v-sheet.theme--light.theme--light>:nth-child(4)'))
    )
    field.click()
    
def wait_load_web(driver, max_sec_wait=30):
    driver.implicitly_wait(0.1)
    t_start = time.time()
    while True:
        time.sleep(0.1)
        t_now = time.time()
        
        if t_now - t_start >= max_sec_wait:
            print('Слишком долгая загрузка сайта. Обновляю странцу.')
            driver.refresh()
            break
        
        try:
            driver.find_element(By.CSS_SELECTOR, '.v-progress-linear__buffer')
            continue
        except:
            break
        
    driver.implicitly_wait(10)

def make_load_dir(name_login, loading_path_part):
    '''Создаю папки в которые сохраняю все данные'''
    
    now = datetime.now()
    dt_string = now.strftime("%Y-%m-%d___%H-%M")
    print("date and time =", dt_string)

    # Создаю папку храенния
    loading_path = f'{loading_path_part}/{name_login} {dt_string}'
    loading_path_all = f'{loading_path}/all_file'
    dowlend_path = loading_path_all

    os.mkdir(loading_path)
    os.mkdir(loading_path_all)
    loading_path_d = loading_path_all
    
    return loading_path, loading_path_d, loading_path_all, dowlend_path

def parsing_all(name_login, date_start, loading_path, loading_path_d, loading_path_all, dowlend_path, list_not_data_for_view):
    '''Основной скрипт для парсинга'''
    
    if date_start is None:
        dict_date_start = {'kpi':date(2021, 1, 1)}
        date_start = dict_date_start[name_login]

    date_finish = date.today()
    date_end = date_finish
    break_flag = False

    status_pars = {'new_driver':True, 'new_log':True, 'set_web_list_targ':True, 'set_iin':True, 'set_100':True}
    cnt_load_file_i = 0
    try_while = 0
    for i in range(1, 200):
        error_list = []
        try_while += 1
        print('try = ', i)
        
        if break_flag:
            break
        
        try:
            t_start_sec = time.time()
        
            if status_pars['new_driver']:
                driver = to_site(loading_path_all)
                status_pars['new_driver'] = False
                
            if len(driver.window_handles) > 1:
                print('Добавлена новая вкладка. Это условие остановки цикла.')
                break_flag = True
                break
                
            if status_pars['new_log']:
                write_login(driver, name_login)
                status_pars['new_log'] = False
                
            if status_pars['set_web_list_targ']:
                to_targer_sheet_1(driver)
                to_targer_sheet_2(driver)
                status_pars['set_web_list_targ'] = False

            print(date_end)
            date_start_srt = date_to_str(date_start)
            date_end_srt   = date_to_str(date_end)

            # Установить дату старта
            butom_date_clear = 'v-input__append-inner'
            field = driver.find_elements(By.CLASS_NAME, butom_date_clear)[0].click()

            butom_date = 'v-text-field__slot'
            field = driver.find_elements(By.CLASS_NAME, butom_date)[0]
            field = field.find_element(By.CSS_SELECTOR, 'input')
            field.send_keys(date_start_srt)
            time.sleep(1)

            # Установка даты окончания 
            butom_date_clear = 'v-input__append-inner'
            field = driver.find_elements(By.CLASS_NAME, butom_date_clear)[1].click()

            butom_date = 'v-text-field__slot'
            field = driver.find_elements(By.CLASS_NAME, butom_date)[1]
            field = field.find_element(By.CSS_SELECTOR, 'input')
            field.send_keys(date_end_srt)
            time.sleep(0.5)    

            # Нажать на поиск
            butom_start_find = '.mb-2.mr-2.mr-md-5.v-btn.v-btn--is-elevated.v-btn--has-bg.theme--light.v-size--default.primary'
            field = driver.find_element(By.CSS_SELECTOR, butom_start_find).click()
            time.sleep(0.5)
            wait_load_web(driver)
            time.sleep(0.5)

            cnt_iter, cnt_row_text = find_cnt_all_row(driver)
            status_pars['set_100'] = False
            
            #new load
            click_load1(driver)
            load_need_value(driver)
            
            for i_iter_list in range(cnt_iter):
                files_before = set(os.listdir(loading_path_all))
                
                if i_iter_list < cnt_load_file_i:
                    click_next_excel_50k(driver)
                    continue
                
                click_load2(driver)
                files_current = set(os.listdir(loading_path_all))
                new_files = files_current - files_before
                if new_files:
                    cnt_load_file_i += 1
                
                if i_iter_list >= (cnt_iter - 1):
                    continue
                click_next_excel_50k(driver)
                try_while = 0
            
            list_not_data_for_view.append([date_end_srt, cnt_row_text])
            if cnt_load_file_i > i_iter_list:
                break

        except Exception as e:
            exc_type, exc_value, exc_traceback = sys.exc_info()
            tb = traceback.extract_tb(exc_traceback)
            str_error = e
            if tb:
                file, line_number, function, text = tb[0]
                str_error = f"{e}\n Error occurred in file {file}, line {line_number}: {text}"
            print(str_error)
            error_list.append(str_error)

            if try_while <= 2:
                driver.refresh()
                status_pars = {'new_driver':False, 'new_log':True, 'set_web_list_targ':True, 'set_iin':True, 'set_100':True}
                time.sleep(5)
            else:
                driver.refresh()
                status_pars = {'new_driver':True, 'new_log':True, 'set_web_list_targ':True, 'set_iin':True, 'set_100':True}
                time.sleep(5)
                try_while = 0
                
                try:
                    driver.quit()
                except:
                    pass


# In[4]:


def save_all_date(loading_path, list_not_data_for_view, loading_path_d):
    '''Сохраняю историю парсинга. Сохраняю Объедененный файл'''
        
    now_str = datetime.now().strftime("%d.%m.%Y")
    pd.DataFrame(list_not_data_for_view).to_excel(f'{loading_path}/Были ли данные {now_str}.xlsx')
    
    list_agg_day_file = []
    for f_name in os.listdir(loading_path_d):
        if f_name[0] == '~':
            continue
        dowlend_file_p = loading_path_d + '//' + f_name
        list_agg_day_file.append(dowlend_file_p)
        
    print()    
    df_concat = pd.DataFrame()
    for f in tqdm(list_agg_day_file):
        df = pd.read_excel(f, dtype={'ИИН/БИН должника':str})
        df_concat = pd.concat([df_concat, df])

    print('Количество строк в объедененном файле =', df_concat.shape[0])    
    
    if df_concat.shape[0] == 0:
        print('В данном АИСОИП, по данным датам нет данных. Переходим к следующему.')
        return None
    
    print('Удаление дублей, Правка №')
    df_final = df_concat.reset_index(False).drop(['index', '№'], axis=1)
    df_final = df_final.drop_duplicates()
    df_final = df_final.reset_index(drop=True).reset_index().rename(columns={'index':'№'})
    df_final['№'] = df_final['№'] + 1
    
    now_str = datetime.now().strftime("%d.%m.%Y")
    df_final.to_excel(f'{loading_path}/{name_login} AIS_OIP {now_str}.xlsx', index=False)
    
    print('Количество строк в объедененном файле без дублей =', df_final.shape[0])
    return df_final

def cleaning_date(df_final):
    '''Очищаю номер года '''
    
    year_recode = {'1202':'2021', 
                   '0202':'2020',
                   '0222':'2022',
                   '0223':'2023',
                   '1120':'2021',
                   '0320':'2023',
                   '0221':'2021',
                   '0218':'2018'
                  }

    def to_date(str_d):
        try:
            split_srt = str_d.split('.')
            if len(split_srt) < 3:
                print(f'Дата {str_d} не может быть преобразована')
                return np.nan

            year_ =  split_srt[-1]
            if '199' == year_:
                print(f'Зануляю дату {str_d}')
                return np.nan

            for year_false, year_true in year_recode.items():
                if year_false == year_:
                    print(f'Заменяю год в дате {str_d} : с {year_true} на {year_false}')
                    return pd.to_datetime(str_d.replace(year_false, year_true), format="%d.%m.%Y")

            return pd.to_datetime(str_d, format="%d.%m.%Y")

        except:
            print(f'Ошибка, строка = "{str_d}" не конвертируется')
            print('Подставлен NaN')
            return np.nan

    df_final = df_final.copy()
    for n_date in ['Дата возбуждения', 'Дата выписки исполнительного документа']:
        print('   ', n_date)
        df_final[n_date] = df_final[n_date].apply(to_date)
        print('+ Все даты верные')
    
    return df_final


# # Парсинг

# In[5]:


for name_login in dict_log_pas.keys():
    print('***', name_login, '***')

    date_start = None
    list_not_data_for_view = []
    
    loading_path, loading_path_d, loading_path_all, dowlend_path = make_load_dir(name_login, loading_path_part)
    
    parsing_all(name_login, date_start, loading_path, loading_path_d, loading_path_all, dowlend_path, list_not_data_for_view)

    df_final = save_all_date(loading_path, list_not_data_for_view, loading_path_d)
    
    if df_final is None:
        continue
        
    df_final = cleaning_date(df_final)
        
    try: 
        time.sleep(5)
        driver.close()
    except:
        pass
    try: 
        time.sleep(2)
        driver.quit()
    except:
        pass
    driver = None


# # Получение данных из Google Sheets и слияние без SQLite

# In[6]:


import string
from pprint import pprint
import pandas as pd
import gspread
from gspread import Cell, Client, Spreadsheet, Worksheet
from gspread.utils import rowcol_to_a1
import requests

SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1DBZnYTpO5vaTh81aELCUuO2DGUSzRSWRHMpPLK8qNg0/edit?usp=drivesdk'

def show_available_worksheets(sh: Spreadsheet):
    worksheets = sh.worksheets()
    for ws in worksheets:
        print("Worksheet with title", repr(ws.title), "and id", ws.id)

def get_dataframe_from_sheet(sh: Spreadsheet, sheet_title: str) -> pd.DataFrame:
    worksheet = sh.worksheet(sheet_title)
    data = worksheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    return df
        
def get_google_df():
    gc: Client = gspread.service_account(r"C:\Users\User\Desktop\py scripts\credentials kpi-api.json")
    sh: Spreadsheet = gc.open_by_url(SPREADSHEET_URL)
    show_available_worksheets(sh)    
    df = get_dataframe_from_sheet(sh, "ММ")
    return df
    
df_google_t = get_google_df()


# In[7]:


def to_int(v):
    v = str(v)
    if len(v.split('.')) > 2 :
        raise f'Ошибка строка {v}'
    
    v = v.replace(' ', '')
    v = v.replace('-', '')
    v = v.replace(',', '.')
    v = v.replace('\xa0', '')
    v = v.split('.')[0]
    if len(v) == 0:
        return 0
    else:
        return int(v)

# Очистка числовых колонок
df_google_t['ОД_claen'] = df_google_t['ОД'].apply(to_int)
columns = ['Сумма выдачи', 'ОД', 'Вознаграждение', 'Пеня', 'Всего задолженность при покупке', 
           'нотариальные услуги', 'сумма оплаты после цессии', 'остаток задолженности', 
           'Сумма оплат', 'госпошина по ауэз суду', 'госпошина по мед суду']
for col in columns:
    print('col', col)
    df_google_t[col] = df_google_t[col].apply(to_int)

df_google_t['госпошина по медеу суду'] = df_google_t['госпошина по мед суду']
df_google_t['ИИН'] = df_google_t['ИИН'].apply(lambda x: '0' * (12 - len(str(x))) + str(x) if str(x) != 'nan' else '')


# In[8]:


# Добавление расчетных полей
df_google_t['госпошина по суду'] = df_google_t['госпошина по ауэз суду'] + df_google_t['госпошина по медеу суду']

df_google_t['f1101'] = df_google_t['Всего задолженность при покупке']
df_google_t['f1102'] = df_google_t['Всего задолженность при покупке'] + df_google_t['нотариальные услуги']
df_google_t['f1103'] = df_google_t['Всего задолженность при покупке'] + df_google_t['госпошина по суду']
df_google_t['f1104'] = df_google_t['Всего задолженность при покупке'] + df_google_t['нотариальные услуги'] + df_google_t['нотариальные услуги']

df_google_t['f1105'] = df_google_t['остаток задолженности']
df_google_t['f1106'] = df_google_t['остаток задолженности'] + df_google_t['нотариальные услуги']
df_google_t['f1107'] = df_google_t['остаток задолженности'] + df_google_t['госпошина по суду']
df_google_t['f1108'] = df_google_t['остаток задолженности'] + df_google_t['нотариальные услуги'] + df_google_t['нотариальные услуги']

df_google_t['f1109'] = df_google_t['Сумма по АИСОИП']
df_google_t['f1110'] = df_google_t['остаток задолженности'] - 685

# Добавляем индекс
df_google_t['index'] = df_google_t.index


# In[9]:


# Подготовка данных АИС ОИП
df_ais = df_final.copy()
df_ais['ИИН/БИН должника'] = df_ais['ИИН/БИН должника'].apply(lambda x: '0' * (12 - len(str(x))) + str(x) if str(x) != 'nan' else '')

# Выбираем нужные колонки
ais_col = ['№', 'Номер исполнительного документа', 'Номер исполнительного производства', 
           'Судебный исполнитель', 'Дата возбуждения', 'Дата выписки исполнительного документа', 
           'Орган выдавший исполнительный документ', 'Тип взыскания', 'Должник',
           'ИИН/БИН должника', 'Статус исполнительного производства', 'Сумма']

df_ais_oip_raw_data = df_ais[ais_col]


# # Слияние данных без SQLite

# In[10]:


# 1. Удаление дубликатов в данных АИС ОИП (аналогично SQL запросу)
def remove_doubles_ais(df_ais):
    """
    Удаляет дубликаты аналогично SQL запросу
    """
    # Сортируем по дате возбуждения в порядке убывания
    df_sorted = df_ais.sort_values('Дата возбуждения', ascending=False)
    
    # Определяем уникальные комбинации
    unique_cols = ['Номер исполнительного документа', 'Номер исполнительного производства',
                   'Судебный исполнитель', 'ИИН/БИН должника', 'Статус исполнительного производства',
                   'Сумма', 'Дата возбуждения']
    
    # Оставляем первую запись для каждой уникальной комбинации
    df_no_doubles = df_sorted.drop_duplicates(subset=unique_cols, keep='first')
    
    return df_no_doubles

# 2. Подготовка данных клиентов
def prepare_clients_data(df_google):
    """
    Подготавливает данные клиентов для слияния
    """
    # Выбираем нужные колонки
    clients_data = df_google[['index', 'ИИН', 'f258', 
                              'f1101', 'f1102', 'f1103', 'f1104',
                              'f1105', 'f1106', 'f1107', 'f1108',
                              'f1109', 'f1110']].copy()
    
    # Переименовываем колонки
    clients_data = clients_data.rename(columns={
        'index': 'eid',
        'ИИН': 'inn',
        'f258': 'f258'
    })
    
    return clients_data

# 3. Функции для слияния данных
def merge_right_data(df_ais_no_doubles, clients_data):
    """
    Слияние по ИИН и номеру производства (первое совпадение)
    """
    # Слияние по inn и f258
    merged = pd.merge(
        df_ais_no_doubles,
        clients_data[['inn', 'f258', 'eid']],
        left_on=['ИИН/БИН должника', 'Номер исполнительного производства'],
        right_on=['inn', 'f258'],
        how='inner'
    )
    
    # Добавляем колонку для приоритета статуса
    merged['priority_id_status'] = merged['Статус исполнительного производства'].apply(
        lambda x: 1 if x == 'На исполнении' else 2
    )
    
    return merged

def merge_by_amount(df_ais_wrong, clients_data):
    """
    Слияние по сумме с допустимой погрешностью
    """
    results = []
    
    for _, row_ais in df_ais_wrong.iterrows():
        amount = row_ais['Сумма']
        inn = row_ais['ИИН/БИН должника']
        
        # Находим клиентов с таким же ИИН
        clients_with_same_inn = clients_data[clients_data['inn'] == inn]
        
        for _, row_client in clients_with_same_inn.iterrows():
            # Проверяем все суммы f1101-f1110
            match_found = False
            for i in range(1, 11):
                col_name = f'f11{i:02d}'
                client_amount = row_client[col_name]
                
                # Пропускаем нулевые суммы
                if client_amount == 0:
                    continue
                
                # Проверяем совпадение с погрешностью 3
                if abs(amount - client_amount) <= 3:
                    match_found = True
                    break
            
            if match_found:
                result = row_ais.copy()
                result['eid'] = row_client['eid']
                result['priority_id_status'] = 1 if result['Статус исполнительного производства'] == 'На исполнении' else 2
                results.append(result)
                break
    
    if results:
        return pd.DataFrame(results)
    else:
        return pd.DataFrame()

def merge_by_single_inn(df_ais_wrong2, clients_data):
    """
    Слияние для клиентов с единственным ИИН
    """
    # Находим клиентов с уникальным ИИН
    inn_counts = clients_data['inn'].value_counts()
    single_inn_clients = inn_counts[inn_counts == 1].index.tolist()
    
    # Фильтруем клиентов с единственным ИИН
    clients_single_inn = clients_data[clients_data['inn'].isin(single_inn_clients)]
    
    # Слияние
    merged = pd.merge(
        df_ais_wrong2,
        clients_single_inn[['inn', 'eid']],
        left_on='ИИН/БИН должника',
        right_on='inn',
        how='inner'
    )
    
    if not merged.empty:
        merged['priority_id_status'] = merged['Статус исполнительного производства'].apply(
            lambda x: 1 if x == 'На исполнении' else 2
        )
    
    return merged

# 4. Основная функция слияния
def merge_data_without_sqlite(df_ais_data, clients_data):
    """
    Основная функция слияния данных без SQLite
    """
    print("1. Удаление дубликатов в данных АИС ОИП...")
    df_ais_no_doubles = remove_doubles_ais(df_ais_data)
    print(f"   После удаления дублей: {len(df_ais_no_doubles)} строк")
    
    print("2. Первое слияние по ИИН и номеру производства...")
    right_data = merge_right_data(df_ais_no_doubles, clients_data)
    print(f"   Найдено совпадений: {len(right_data)}")
    
    # Находим строки, которые не попали в первое слияние
    print("3. Поиск неподтянутых строк...")
    merged_indexes = set(right_data.index)
    all_indexes = set(df_ais_no_doubles.index)
    wrong_indexes = all_indexes - merged_indexes
    
    df_ais_wrong = df_ais_no_doubles.loc[list(wrong_indexes)].copy()
    print(f"   Неподтянуто строк: {len(df_ais_wrong)}")
    
    print("4. Второе слияние по сумме с погрешностью...")
    right_data2 = merge_by_amount(df_ais_wrong, clients_data)
    print(f"   Найдено совпадений по сумме: {len(right_data2)}")
    
    # Находим оставшиеся неподтянутые строки
    if not right_data2.empty:
        merged_indexes2 = set(right_data2.index)
        wrong_indexes2 = set(df_ais_wrong.index) - merged_indexes2
        df_ais_wrong2 = df_ais_wrong.loc[list(wrong_indexes2)].copy()
    else:
        df_ais_wrong2 = df_ais_wrong.copy()
    
    print(f"   Осталось неподтянуто: {len(df_ais_wrong2)}")
    
    print("5. Третье слияние для клиентов с единственным ИИН...")
    right_data3 = merge_by_single_inn(df_ais_wrong2, clients_data)
    print(f"   Найдено совпадений для уникальных ИИН: {len(right_data3)}")
    
    # Объединяем все найденные совпадения
    print("6. Объединение всех результатов...")
    all_right_data = []
    
    if not right_data.empty:
        all_right_data.append(right_data)
    if not right_data2.empty:
        all_right_data.append(right_data2)
    if not right_data3.empty:
        all_right_data.append(right_data3)
    
    if all_right_data:
        combined = pd.concat(all_right_data, ignore_index=True)
        
        # Удаляем дубликаты по eid, оставляя самые приоритетные
        combined_sorted = combined.sort_values(['eid', 'priority_id_status', 'Дата возбуждения'], 
                                              ascending=[True, True, False])
        final_data = combined_sorted.drop_duplicates('eid', keep='first')
    else:
        final_data = pd.DataFrame()
    
    # Находим окончательно неподтянутые строки
    print("7. Определение окончательно неподтянутых строк...")
    if not final_data.empty:
        wrong_final = df_ais_no_doubles[
            ~df_ais_no_doubles['Номер исполнительного производства'].isin(final_data['Номер исполнительного производства'])
        ].copy()
    else:
        wrong_final = df_ais_no_doubles.copy()
    
    # Добавляем статус "На исполнении" для фильтрации
    wrong_final_active = wrong_final[wrong_final['Статус исполнительного производства'] == 'На исполнении'].copy()
    
    print(f"   Итого подтянуто: {len(final_data) if not final_data.empty else 0}")
    print(f"   Итого не подтянуто (в работе): {len(wrong_final_active)}")
    
    return final_data, wrong_final_active


# In[11]:


# Подготавливаем данные клиентов
clients_data = prepare_clients_data(df_google_t)

# Выполняем слияние
final_data, wrong_data_active = merge_data_without_sqlite(df_ais_oip_raw_data, clients_data)


# In[12]:


# Проверки
if not final_data.empty:
    print("Проверка результатов:")
    print(f"1. Кол-во уникальных eid: {final_data['eid'].nunique()}")
    print(f"2. Кол-во строк в результатах: {len(final_data)}")
    print(f"3. Пустые eid: {final_data['eid'].isna().sum()}")
    
    # Проверка на дубликаты eid
    eid_counts = final_data['eid'].value_counts()
    duplicates = eid_counts[eid_counts > 1]
    print(f"4. Дублей eid: {len(duplicates)}")
    
    # Проверка незавершенных дел
    is_finish = (df_google_t['статус кредита'] == 'погашен') | (df_google_t['статус кредита'] == 'обратный выкуп')
    active_clients = df_google_t.loc[~is_finish, 'index']
    matched_active = active_clients.isin(final_data['eid'])
    print(f"5. Не подтянуто из незавершенных: {(~matched_active).sum()}")
else:
    print("Нет данных в final_data")


# In[13]:


# Подготовка финальной таблицы
final_table = df_google_t[['index', 'ИИН']].copy()
final_table = final_table.rename(columns={'index': 'Номер_строки_google_t'})
final_table['Номер_строки_google_t'] = final_table['Номер_строки_google_t'].astype('Int64')

if not final_data.empty:
    # Присоединяем данные из слияния
    merged = pd.merge(
        final_table,
        final_data[['eid'] + ais_col[1:]],  # все колонки кроме '№'
        left_on='Номер_строки_google_t',
        right_on='eid',
        how='left'
    )
    
    # Переименовываем eid в Уникальный идентификатор займа
    merged = merged.rename(columns={'eid': 'Уникальный идентификатор займа'})
else:
    # Если нет совпадений, создаем пустые колонки
    merged = final_table.copy()
    for col in ais_col[1:]:
        merged[col] = None
    merged['Уникальный идентификатор займа'] = None

# Сортировка
merged = merged.sort_values('Номер_строки_google_t')

# Функция для создания уникального номера
def uid_to_string(uid, first_num=2, cnt_num=6):
    suid = str(uid)
    cnt_zero = cnt_num - 1 - len(suid)
    new_suid = str(first_num) + "0" * cnt_zero + suid
    return new_suid

merged['уникальный номер'] = merged['Номер_строки_google_t'] + 1
merged['уникальный номер'] = merged['уникальный номер'].apply(uid_to_string)

# Если есть колонка Уникальный идентификатор займа, удаляем ее
if 'Уникальный идентификатор займа' in merged.columns:
    merged = merged.drop('Уникальный идентификатор займа', axis=1)


# In[14]:


# Сохранение результатов
now_str = datetime.now().strftime("%d.%m.%Y")
merged.to_excel(f'{loading_path}/{name_login} AIS_OIP_merge_google_table {now_str}.xlsx', index=False)
print(f"Сохранен файл: {loading_path}/{name_login} AIS_OIP_merge_google_table {now_str}.xlsx")

# Сохранение неподтянутых производств
if not wrong_data_active.empty:
    wrong_data_active.to_excel(f'{loading_path}/{name_login} Производства не подтянулись.xlsx', index=False)
    print(f"Производств в работе не найдено: {len(wrong_data_active)}")
else:
    print("Все производства в работе найдены!")


# In[15]:


# Убираю из проверенных отмененных (не отменённые статусы)
def clean_cancel_list(new_clean_df, path_load_finish):
    '''
    Очитаю эскль в котором хранятся данные по отменам которыне не нужно проверять. 
    '''
    # Преобразуем дату если нужно
    if 'Дата возбуждения' in new_clean_df.columns:
        try:
            new_clean_df['Дата возбуждения'] = pd.to_datetime(new_clean_df['Дата возбуждения'], format='%d.%m.%Y')
            new_clean_df = new_clean_df.sort_values(by=['Дата возбуждения'], ascending=False)
        except:
            pass
    
    # Создаем уникальный идентификатор
    new_clean_df['ИИН_and_id_Испол'] = new_clean_df['ИИН'] + '_' + new_clean_df['Номер исполнительного производства']
    
    path_check_iin = Path(path_load_finish).joinpath('Проверенные ранее.xlsx')
    
    if os.path.exists(path_check_iin):
        df_old_ = pd.read_excel(path_check_iin, dtype='str')
        
        if 'ИИН_and_id_Испол' not in df_old_.columns:
            df_old_['ИИН_and_id_Испол'] = df_old_['ИИН'] + '_' + df_old_['Номер исполнительного производства']
        
        # Удаляем отмененные производства
        if 'Статус исполнительного производства' in new_clean_df.columns:
            not_canceled_mask = new_clean_df['Статус исполнительного производства'] != 'Окончено'
            active_productions = new_clean_df.loc[not_canceled_mask, 'ИИН_and_id_Испол']
        else:
            active_productions = new_clean_df['ИИН_and_id_Испол']
        
        df_new_check = df_old_[~df_old_['ИИН_and_id_Испол'].isin(active_productions)].copy()
        df_new_check.to_excel(path_check_iin, index=False)
        
        print(f'Удалено из проверенных {len(df_old_) - len(df_new_check)}')
    else:
        print(f'Файл {path_check_iin} не найден')

# Вызываем функцию очистки
clean_cancel_list(merged, path_load_finish)


# In[16]:


# Создание дополнительного файла с форматами
import re
import glob

def get_last_cancel_path(load_path):
    date_pattern = re.compile(r"(\d{4}-\d{2}-\d{2})___\d{2}-\d{2}")
    latest_date = None
    latest_folder = None

    for folder in os.listdir(load_path):
        folder_path = os.path.join(load_path, folder)
        if not os.path.isdir(folder_path):
            continue
        match = date_pattern.match(folder)
        if match:
            folder_date = datetime.strptime(match.group(1), "%Y-%m-%d")
            if latest_date is None or folder_date > latest_date:
                latest_date = folder_date
                latest_folder = folder

    print(f"Самая последняя папка: {latest_folder}")
    return latest_folder

latest_folder_cancel = get_last_cancel_path(path_load_finish)

# Загрузка данных об отменах
if latest_folder_cancel:
    latest_folder_path = os.path.join(path_load_finish, latest_folder_cancel)
    file_pattern = os.path.join(latest_folder_path, "*Итог_union*")
    matching_files = glob.glob(file_pattern)
    
    if matching_files:
        target_file = matching_files[0]
        print(f"Найденный файл: {target_file}")
        df_last_cancel = pd.read_excel(target_file)
    else:
        print("Файл с 'Итог_union' не найден.")
        df_last_cancel = pd.DataFrame()
else:
    df_last_cancel = pd.DataFrame()

# Загрузка проверенных ранее
check_later = os.path.join(path_load_finish, 'Проверенные ранее.xlsx')
if os.path.exists(check_later):
    df_check_later = pd.read_excel(check_later)
else:
    df_check_later = pd.DataFrame()

# Объединение данных об отменах
if not df_last_cancel.empty:
    df_last_cancel = df_last_cancel.rename(columns={
        "inn": 'ИИН', 
        "f258": 'Номер исполнительного производства', 
        "Класс 1": "Класс 1",
        "Дата 1": "Дата 1"
    })

if not df_check_later.empty and "Дата 1" in df_check_later.columns:
    df_check_later["Дата 1"] = pd.to_datetime(df_check_later["Дата 1"])

if not df_last_cancel.empty and "Дата 1" in df_last_cancel.columns:
    df_last_cancel["Дата 1"] = pd.to_datetime(df_last_cancel["Дата 1"], format="%d.%m.%Y", errors='coerce')

# Объединяем все отмены
df_cancel = pd.DataFrame()
if not df_check_later.empty:
    df_cancel = pd.concat([df_cancel, df_check_later[['ИИН', 'Номер исполнительного производства', 'Класс 1', "Дата 1"]]])
if not df_last_cancel.empty:
    df_cancel = pd.concat([df_cancel, df_last_cancel[['ИИН', 'Номер исполнительного производства', 'Класс 1', "Дата 1"]]])

if not df_cancel.empty:
    df_cancel = df_cancel.dropna()
    mask = (df_cancel["Класс 1"].notna()) & (df_cancel["Класс 1"] != "")
    df_cancel = df_cancel.loc[mask]
    df_cancel = df_cancel.rename(columns={
        "Класс 1": "Статус исполнительного документа", 
        "Дата 1": "Дата Постановления о прекращении ИП"
    })
    df_cancel = df_cancel.drop_duplicates(["ИИН", "Номер исполнительного производства"])


# In[17]:


# Создание нового файла для импорта
col_target = [
    "Уникальный номер сделки", "Статус АИС ОИП", "ЧСИ", 
    "Дата возбуждения судебного производства", "Сумма иска", 
    "Номер исполнительного производства", "Номер исполнительного документа", 
    "Статус исполнительного документа", "Дата выписки исполнительного документа", 
    "Орган выдавший исполнительный документ", "Дата Постановления о прекращении ИП"
]

# Подготовка данных для импорта
final_table_col_rename = merged.copy()

# Переименование колонок
rename_columns = {
    "уникальный номер": "Уникальный номер сделки",
    "Статус исполнительного производства": "Статус АИС ОИП",
    "Судебный исполнитель": "ЧСИ",
    "Дата возбуждения": "Дата возбуждения судебного производства",
    "Сумма": "Сумма иска",
    "Номер исполнительного производства": "Номер исполнительного производства",
    "Номер исполнительного документа": "Номер исполнительного документа",
    "Дата выписки исполнительного документа": "Дата выписки исполнительного документа",
    "Орган выдавший исполнительный документ": "Орган выдавший исполнительный документ"
}

# Применяем переименование только для существующих колонок
for old_name, new_name in rename_columns.items():
    if old_name in final_table_col_rename.columns:
        final_table_col_rename = final_table_col_rename.rename(columns={old_name: new_name})

# Объединяем с данными об отменах
if not df_cancel.empty:
    # Приводим ИИН к одному типу
    final_table_col_rename["ИИН"] = pd.to_numeric(final_table_col_rename["ИИН"], errors='coerce')
    df_cancel["ИИН"] = pd.to_numeric(df_cancel["ИИН"], errors='coerce')
    
    # Объединяем
    final_table_col_rename = pd.merge(
        final_table_col_rename, 
        df_cancel[["ИИН", "Номер исполнительного производства", "Статус исполнительного документа", "Дата Постановления о прекращении ИП"]],
        on=["ИИН", "Номер исполнительного производства"], 
        how='left'
    )
else:
    # Добавляем пустые колонки если нет данных об отменах
    final_table_col_rename["Статус исполнительного документа"] = None
    final_table_col_rename["Дата Постановления о прекращении ИП"] = None

# Выбираем только нужные колонки
existing_cols = [col for col in col_target if col in final_table_col_rename.columns]
final_table_import = final_table_col_rename[existing_cols].copy()


# In[18]:


# Сохранение файла для импорта
now_str = datetime.now().strftime("%d.%m.%Y")
path_format1 = f'{loading_path}/AIS_OIP_Import.xlsx'
final_table_import.to_excel(path_format1, index=False)
print(f"Создан файл для импорта: {path_format1}")

# Применение форматирования из шаблона
template_path = f'{loading_path_part}/Шаблон1.xlsx'

if os.path.exists(template_path):
    try:
        import win32com.client as win32
        
        def apply_format_and_header_from_template(template_path, target_path):
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            
            try:
                template_wb = excel.Workbooks.Open(template_path)
                target_wb = excel.Workbooks.Open(target_path)
                
                template_ws = template_wb.Sheets(1)
                target_ws = target_wb.Sheets(1)
                
                # Копируем формат заголовка
                template_header_range = template_ws.Rows(1)
                target_header_range = target_ws.Rows(1)
                template_header_range.Copy()
                target_header_range.PasteSpecial(Paste=-4104)
                
                # Применяем формат второй строки
                template_second_row_format = template_ws.Rows(2)
                target_last_row = target_ws.UsedRange.Rows.Count
                target_body_range = target_ws.Rows(f"2:{target_last_row}")
                template_second_row_format.Copy()
                target_body_range.PasteSpecial(Paste=-4122)
                
                # Копируем ширину столбцов
                template_used_columns = template_ws.UsedRange.Columns.Count
                for col in range(1, template_used_columns + 1):
                    template_column_width = template_ws.Columns(col).ColumnWidth
                    target_ws.Columns(col).ColumnWidth = template_column_width
                
                target_wb.Save()
                print("Форматирование применено успешно")
                
            finally:
                template_wb.Close(SaveChanges=False)
                target_wb.Close(SaveChanges=True)
                excel.Quit()
        
        apply_format_and_header_from_template(template_path, path_format1)
    except Exception as e:
        print(f"Ошибка при применении форматирования: {e}")
else:
    print(f"Шаблон не найден: {template_path}")


# In[19]:


# Копирование файла в папку CRM
path_format1 = Path(path_format1)
path_crm = Path(path_crm)
path_crm_final = path_crm / path_format1.name

try:
    shutil.copy2(path_format1, path_crm_final)
    print(f"Файл успешно скопирован в: {path_crm_final}")
except Exception as e:
    print(f"Ошибка при копировании файла: {e}")


# In[20]:


print("=" * 50)
print("ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО!")
print("=" * 50)
print(f"1. Данные скачаны с АИС ОИП: {len(df_final)} строк")
print(f"2. Данные загружены из Google Sheets: {len(df_google_t)} строк")
print(f"3. Совпадений найдено: {len(final_data) if not final_data.empty else 0}")
print(f"4. Производств не подтянулось: {len(wrong_data_active)}")
print(f"5. Файл для импорта создан: {path_format1}")
print("=" * 50)
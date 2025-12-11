#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# # Ввод данных

# Данные аутентификации:
# - логин;
# - пароль;
# - начало названия организации (в боковой левой панели)

# In[123]:


USER_AUTH = '210540032049'
USER_PASSWORD = 'Qazaq123456'
USER_NAME_BEGIN = 'ТОВАРИЩЕСТВО С ОГРАНИЧЕННОЙ ОТ'



# # настройки

# In[124]:


PATH_TO_EXCELFILE = 'load_суд/'  # Путь - пример: 'C:/orders/'
FILE_NAME = '' # имя файла

continue_parsing = False # Продолжить скачивание? (Fales - для скачивания сначала.)


# # Загрузка и импорт необходимых библиотек

# In[125]:


# !python.exe -m pip install --upgrade pip
# !pip3 install -U selenium
# !pip3 install -U xlsxwriter
# !pip3 install -U pandas
# !pip3 install bs4


# In[126]:


from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException 
from pathlib import Path
from selenium.webdriver.common.action_chains import ActionChains

import pandas as pd
import datetime
from typing import Union
from time import sleep
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, WebDriverException
import time
import logging

from time import sleep
from selenium.common.exceptions import TimeoutException
import logging
import os
import requests
import json
import pickle
import random
import time
import logging
import socket
import ssl
from urllib.parse import urlsplit
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError



from bs4 import BeautifulSoup
import re
import numpy as np
pd.set_option('display.max_columns', None)


# In[127]:


sleep(2)


# In[128]:


# Логи

import os
import logging

# Указываем путь для сохранения файла
log_file_path = 'log.txt'

# Получаем логгер
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Добавляем обработчик только если он ещё не добавлен
if not logger.handlers:
    # Файловый обработчик
    file_handler = logging.FileHandler(log_file_path, mode='a', encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(funcName)s - %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

# Пример использования
logger.info("Логирование работает — проверка записи в файл")


# # Парсинг сайта

# ## Постоянные данные для загрузки сайта

# In[129]:


LOADING_TIME = 60    # Время ожидания загрузки страниц, в секундах
NUMBER_OF_PAGE_LOAD_ATTEMPTS = 3    # Количество попыток загрузки. Если страница не прогрузилась, то запрос повторяется.  
NUMBER_OF_BROWSER_REOPENINGS = 100    # Количество повторных открытий браузера


# ## Подготовка функций

# ### Парсинг сайта

# Вызов исключения в случае отсутствия доступа к странице

# In[130]:


def decorator_throw_exception(func):
    def wrapper(*args, **kwargs):
        list_params = [*args]
        dict_arg = dict(**kwargs)
        if 'iteration' in dict_arg:
            if dict_arg['iteration'] > NUMBER_OF_PAGE_LOAD_ATTEMPTS:
                raise WebDriverException(f'Доступ к сайту прерван. Последняя удачно загруженная страница: {list_params[0].current_url}')
        sleep(0.1)
        func(*args, **kwargs)
    return wrapper


# Вход на русскую версию сайта

# In[131]:


# @decorator_throw_exception
def login_to_site(driver, iteration=0):

    url = 'https://office.sud.kz/index.xhtml'

    main_page(driver, url)

    driver.refresh()
    sleep(1)
    if 'судебный кабинет' not in driver.title.lower():
        change_lang_on_ru(driver)

    # Поле пароля по placeholder

    field_password = driver.find_element(
        By.CSS_SELECTOR, 'input[placeholder="Пароль"]'
    )

    field_login = WebDriverWait(driver, LOADING_TIME).until(
        EC.visibility_of_element_located((
            By.XPATH,
            "//input[@placeholder='ИИН/БИН' or @id='j_idt94:auth:xin']"
        ))
    )


    come_in = driver.find_element(
        By.CSS_SELECTOR, 'input.button-primary[type="submit"]'
    )



    auth = USER_AUTH
    password = USER_PASSWORD

    sleep(0.5)

    field_login.clear()
    field_login.send_keys(auth)
    field_password.clear()
    field_password.send_keys(password)

    come_in.click()
    sleep(1)
#     try:
#         WebDriverWait(driver, LOADING_TIME).until(
#             EC.presence_of_element_located((By.CSS_SELECTOR, ".flex.flex-ai-center.flex-jc-center.flex-col"))
#         )

#     except TimeoutException:
#         iteration +=1
#         login_to_site(driver, iteration=iteration)


# Вход на главную страницу сайта

# In[132]:


@decorator_throw_exception
def main_page(driver, url:str, iteration=0):

    driver.get(url)
    sleep(0.5)

    try:
        pass
        # WebDriverWait(driver, LOADING_TIME).until(EC.element_to_be_clickable((By.CLASS_NAME, 'logo-block')))
    except TimeoutException:
        iteration +=1
        main_page(driver, url, iteration=iteration)


# Изменяет язык сайта на русский

# In[133]:



def change_lang_on_ru(driver, attempts=10, timeout=10):
    for i in range(attempts):
        try:
            btn = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "РУС"))
            )
            btn.click()
            WebDriverWait(driver, timeout).until(
                EC.title_contains("Судебный кабинет")
            )
            return
        except TimeoutException:
            sleep(1)
    raise TimeoutException("Не удалось переключить язык на русский за 5 попыток")


# переход к странице дел

# In[135]:



# @decorator_throw_exception
# def go_to_workcab(driver, iteration=0):
#     url_cab = 'https://office.sud.kz/new/form/cases/mycases.xhtml'
#     driver.get(url_cab)
#     sleep(0.5)
#     try:
#         WebDriverWait(driver, LOADING_TIME).until(
#             EC.presence_of_element_located((By.CLASS_NAME, 'case-item-container'))
#         )
#     except TimeoutException:
#         iteration +=1
#         go_to_workcab(driver, iteration=iteration)



"""
def go_to_workcab(driver, max_attempts: int = 5, loading_time: int = LOADING_TIME):
    url_cab = "https://office.sud.kz/form/cases/mycases.xhtml" # <новая ссылка.  'https://office.sud.kz/new/form/cases/mycases.xhtml' <было
    last_url = None

    for attempt in range(1, max_attempts + 1):
        driver.get(url_cab)
        sleep(0.5)

        try:
            WebDriverWait(driver, loading_time).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'case-item-container'))
            )
            logger.info("Отработал go_to_workcab корректно")

            return  # успех: элемент найден, выходим из функции
        except TimeoutException:
            last_url = getattr(driver, 'current_url', url_cab)
            if attempt == max_attempts:
                raise WebDriverException(
                    f'Доступ к сайту прерван. Последняя удачно загруженная страница: {last_url}'
                )
            # небольшая пауза перед следующей попыткой
            sleep(0.1)
"""
def _visible_case_cards(driver):
    cards = driver.find_elements(By.CSS_SELECTOR, ".case-item-container")
    # фильтрация только видимых
    return [c for c in cards if c.is_displayed() and c.size.get('height', 0) > 0 and c.size.get('width', 0) > 0]

def go_to_workcab(driver, max_attempts: int = 5, loading_time: int = LOADING_TIME):
    # единый URL
    url_cab = "https://office.sud.kz/form/cases/mycases.xhtml"

    for attempt in range(1, max_attempts + 1):
        driver.get(url_cab)

        try:
            # 1) Если уже есть видимые карточки — успех
            WebDriverWait(driver, loading_time).until(lambda d: len(_visible_case_cards(d)) > 0)
            logger.info("go_to_workcab: карточки видимы (через GET)")
            return
        except TimeoutException:
            pass  # пробуем кликом по меню

        try:
            # 2) Пробуем кликнуть по пункту меню «Мои дела»
            # (варианты локаторов в зависимости от верстки)
            menu = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
                By.XPATH, "//a[.//text()[contains(., 'Мои дела')] or contains(., 'Мои дела')]"
            )))
            menu.click()

            # 3) Ждём, пока появятся ВИДИМЫЕ карточки
            WebDriverWait(driver, loading_time).until(lambda d: len(_visible_case_cards(d)) > 0)
            logger.info("go_to_workcab: карточки видимы (после клика по меню)")
            return
        except TimeoutException:
            if attempt == max_attempts:
                last_url = getattr(driver, 'current_url', url_cab)
                raise WebDriverException(
                    f'Не удалось открыть «Мои дела». Последний URL: {last_url}'
                )
            sleep(0.5)


# Получение количества страниц

# In[136]:


def get_count_pages(driver) -> int:
    div_list_pages = driver.find_element(By.CLASS_NAME,'list-pages')
    list_pages_links = div_list_pages.find_elements(By.TAG_NAME, "a")
    return int(list_pages_links[-2].text)


# Переходит на указанную страницу

# In[137]:





def go_to_page(driver, count_page: int, *, max_retries: int = 5) -> None:
    """
    Переходит на страницу пагинации `count_page`.
    Делает до `max_retries` попыток при проблемах с загрузкой/ожиданиями.
    Бросает исключение на неуспехе.
    """
    url = "https://office.sud.kz/form/cases/mycases.xhtml"
    last_error = None

    for attempt in range(1, max_retries + 1):
        try:
            # 1) открыть страницу с пагинацией
            driver.get(url)

            # если что-то не успело прогрузиться — чуть подождём
            sleep(0.2)

            # 2) найти любую ссылку пагинатора с RichFaces.ajax
            sample = WebDriverWait(driver, LOADING_TIME).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a[onclick*="RichFaces.ajax"]'))
            )
            onclick = sample.get_attribute("onclick") or ""

            # 3) вытащить widgetId из этого onclick
            m = re.search(r'RichFaces\.ajax\("([^"]+)"', onclick)
            if not m:
                raise TimeoutException("Не удалось извлечь widgetId из onclick.")

            widget_id = m.group(1)

            # 4) собрать новый JS-вызов с нужным count_page
            js_call = (
                f'RichFaces.ajax("{widget_id}", event, '
                f'{{"incId":"1","parameters":{{"thisPage":"{count_page}"}}}});'
            )
            # (не забываем отбросить `return false;` — здесь он и не нужен)

            # 5) выполнить перелистыватель
            driver.execute_script(js_call)

            # 6) дождаться, что в пагинаторе подсветится именно count_page
            WebDriverWait(driver, LOADING_TIME).until(
                EC.text_to_be_present_in_element(
                    (By.CSS_SELECTOR, ".current"),
                    str(count_page)
                )
            )
            # Успех — выходим
            return

        except (TimeoutException, StaleElementReferenceException, JavascriptException, NoSuchElementException) as e:
            last_error = e
            if attempt >= max_retries:
                # на последней попытке — пробрасываем
                raise
            # небольшая экспоненциальная пауза перед повтором
            sleep(0.5 * attempt)

    # на всякий случай, если цикл закончился без return/raise
    if last_error:
        raise last_error


# Получение количества дел на странице

# In[138]:


"case-item-container flex flex-jc-start hovered mb-20" == "case-item-container flex flex-jc-start hovered mb-20"


# In[139]:


# def count_cases_on_page(driver) -> int:
#     return len(driver.find_elements(By.CLASS_NAME,'list-folders-item'))

def count_cases_on_page(driver) -> int:
    return len(driver.find_elements(By.CLASS_NAME, "case-item-container flex flex-jc-start hovered mb-20".replace(" ", ".")))


# Вход в дело




# Проблемы с выходом из дела.
def come_to_case(driver, count_case: int, iteration: int = 0):
    try:
        logger.info(f"Поиск дел (попытка {iteration}, count_case={count_case})")
        # Ждём появления всех дел
        cases = WebDriverWait(driver, LOADING_TIME).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "case-item-container"))
        )

        if count_case >= len(cases):
            raise IndexError(f"count_case={count_case} вне диапазона: найдено {len(cases)} дел.")

        logger.info(f"Клик по делу #{count_case} — {cases[count_case].text.strip().splitlines()[0]}")

        # Скролл + клик по нужному делу
        ActionChains(driver).move_to_element(cases[count_case]).click().perform()

        # Ждём, когда откроется страница дела (по наличию класса)
        WebDriverWait(driver, LOADING_TIME).until(
            EC.presence_of_element_located((By.CLASS_NAME, "form-static-info"))
        )

        logger.info(f"Дело #{count_case} успешно открыто.")

    except TimeoutException:
        logger.warning(f"Не дождались открытия дела #{count_case} — повторная попытка {iteration + 1}")
        if iteration < 2:
            time.sleep(0.1)
            come_to_case(driver, count_case=count_case, iteration=iteration + 1)
        else:
            logger.error(f"Повторные попытки не помогли — дело #{count_case} не открылось.")
            raise

    except Exception as e:
        logger.exception(f"Ошибка при попытке открытия дела #{count_case}: {e}")
        raise


# In[ ]:





# Получение номера дела

# In[141]:


def get_case_number(driver) -> str:
    return driver.find_element(By.CLASS_NAME,'form-static-info').text


# In[142]:


def get_case_number_by_index(driver, count_case: int, iteration: int = 0) -> str:
    try:
        cases = WebDriverWait(driver, LOADING_TIME).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "case-item-container"))
        )
        if count_case >= len(cases):
            raise IndexError(f"Дело с индексом {count_case} не найдено. Найдено только {len(cases)} элементов.")
        case_element = cases[count_case]
        h3 = case_element.find_element(By.TAG_NAME, "h3")
        return h3.text.strip()

    except Exception as e:
        logging.error(f"Попытка {iteration}: Ошибка при получении номера дела: {e}")
        if iteration >= NUMBER_OF_PAGE_LOAD_ATTEMPTS:
            return ""
        sleep(0.5)
        return get_case_number_by_index(driver, count_case, iteration=iteration+1)


# In[ ]:





# Определяем тип окна статусов дела. `collapseDynamicReview` - если поле имеет название "ДИНАМИКА ХОДА РАССМОТРЕНИЯ ДЕЛА", `collapseHistory` - если поле имеет название "СТАТУСЫ". Если это другие типы статусов или статусы не найдены, то сообщаем об этом и переходим к другому делу, дело в выгрузку не попадет

# In[143]:


def type_of_case_dynamics(driver, case_number:str) -> str:
    if len(driver.find_elements(By.ID,'collapseDynamicReview')) > 0:
        return 'collapseDynamicReview'
    elif len(driver.find_elements(By.ID,'collapseHistory')) > 0:
        return 'collapseHistory'
    else:
        return ''


# Получение списка элементов записей дела

# In[144]:


def get_dynamics_case(driver, type_dynamics:str) -> list:
    block_consideration_case = driver.find_element(By.ID,type_dynamics)
    return block_consideration_case.find_elements(By.CLASS_NAME,'well')


# Получение даты записи

# In[145]:


def get_date_case_record(case_record, type_dynamics:str) -> str:
    if type_dynamics == 'collapseDynamicReview':
        return case_record.find_elements(By.TAG_NAME ,'p')[1].get_attribute('innerHTML')
    elif type_dynamics == 'collapseHistory':
        return case_record.find_element(By.TAG_NAME ,'span').get_attribute('innerHTML')


# Получение текста записи

# In[146]:


def get_text_case_record(case_record, type_dynamics:str) -> str:
    if type_dynamics == 'collapseDynamicReview':
        return case_record.find_elements(By.TAG_NAME ,'div')[1].get_attribute('innerHTML')
    elif type_dynamics == 'collapseHistory':
        return case_record.get_attribute('innerHTML')


# Получение значения из элемента

# In[147]:


# def get_value_from_element(driver, element_type:By, element_name:str) -> str:
#     return driver.find_element(element_type,element_name).get_attribute('value')

def get_value_from_element(driver, element_type: By, element_name: str) -> str:
    container = driver.find_element(element_type, element_name)
    input_el = container.find_element(By.TAG_NAME, "input")
    return input_el.get_attribute("value")


# Получение значение колонки строки

# In[148]:



def get_column_value(row_element, column_index, lower=False) -> str:
    cells = row_element.find_elements(By.CLASS_NAME, 'cell')
    if column_index >= len(cells):
        return ""
    value = cells[column_index].text.strip()
    return value.lower() if lower else value


# Получить строки таблицы

# In[149]:


# def get_table_rows(table) -> list:
#     return table.find_elements(By.TAG_NAME,'tr')

def get_table_rows(table) -> list:
    return table.find_elements(By.CLASS_NAME, 'row')


# ### Обработка данных

# Имена колонок для DF

# In[150]:


LIST_COL = [
    'list_defendants', # dict(defendant_inn, defendant_fio)
    'plaintiff_representative',
    'number_case',
    'location',
    'judicial_authority',
    'case_category',
    'amount_of_claim',
    'amount_of_state_duties',
    'date_of_departure',
    'date_of_rejection',
    'cause_of_rejection',
    'date_of_registration',
    'name_judge',
    'date_of_return',
    'date_of_agreement',
    'date_of_simplification',
    'date_of_first_instance',
    'date_court_order',
    'date_refusal_of_summary_proceedings'
]


# Связь имен колонок с текстом записи

# In[ ]:


DICT_SEARCH_NAME_COL = {
    'Исковое заявление отправлено': 'date_of_departure',
    'отклонено': 'date_of_rejection',
    'зарегистрировано': 'date_of_registration',
    'вынесено определение о возврате искового заявления': 'date_of_return',
    'вынесено определение об утверждении соглашения об урегулировании спора': 'date_of_agreement',
    'вынесено определение о рассмотрении дела в порядке упрощенного производства': 'date_of_simplification',
    'вынесено решение первой инстанции': 'date_of_first_instance',
    'вынесено судебный приказ': 'date_court_order',
    'об отмене решения в порядке упрощенного (письменного) производства': 'date_refusal_of_summary_proceedings',
    'Заявление успешно отправлено': 'date_of_departure',
    'Иск отправлено': 'date_of_departure',
}


# Первичная структура для подготовки к записи в DF

# In[152]:


def create_dict_for_table() -> dict:
    return dict(zip(LIST_COL,['' for i in range(len(LIST_COL))]))    


# Выбор имени колонки по тексту записи

# In[153]:


def select_col_by_text(text:str) -> Union[str, None]:
    for part_string, name_col in DICT_SEARCH_NAME_COL.items():
        if part_string.lower() in text.lower():
            return name_col


# Поиск фамилии судьи в тексте записи

# In[154]:


def search_name_judge(text:str, name_judge: str) -> str:
    index_judge = text.rfind('Судья –')
    if index_judge >= 0:
        start_index_name_judge = index_judge + 8
        end_index_name_judge = start_index_name_judge + text[start_index_name_judge:].find(' ')
        return text[start_index_name_judge:end_index_name_judge]
    else:
        return name_judge


# Пояснения по отказу заявления

# In[155]:


def text_cuase_of_rejection(text:str) -> str:
    index_cause = text.rfind('Причина: ')
    if index_cause >= 0:
        start_index_cause = index_cause + 9
        return text[start_index_cause:].replace('Телеграм бот: <a href="https://t.me/smartsot_bot" target="_blank">https://t.me/smartsot_bot</a>', '').strip()
    else:
        return ''


# Запись данных из истории дела

# In[156]:


def recording_information_from_records(dynamics_case:list, dict_case_for_table:dict, type_dynamics:str) -> None: 
    name_judge = ''                    
    for case_record in dynamics_case: # проход по элементам записей
        date_case_record = get_date_case_record(case_record, type_dynamics) # получение даты записи
        date_for_table = date_case_record[:10] # убираем время
        text_case_record = get_text_case_record(case_record, type_dynamics) # получение текста записи

        name_judge = search_name_judge(text_case_record, name_judge)

        name_col_in_table = select_col_by_text(text_case_record)
        if name_col_in_table == None:
            continue

        # пояснения при отклонении заявления
        if name_col_in_table == 'date_of_rejection':
            dict_case_for_table['cause_of_rejection'] = text_cuase_of_rejection(text_case_record)

        # проверка на наличие даты у колонки
        if dict_case_for_table[name_col_in_table] != '':
            if name_col_in_table == 'date_of_first_instance':
                continue
            dict_case_for_table[name_col_in_table] = f'{dict_case_for_table[name_col_in_table]}, {date_for_table}'
            continue

        dict_case_for_table[name_col_in_table] = date_for_table

    dict_case_for_table['name_judge'] = name_judge


# Запись дополнительной информации по делу

# In[157]:



def filling_in_additional_information(driver, dict_case_for_table:dict) -> None:
    try:
        dict_case_for_table['location'] = get_value_from_element(driver, By.CLASS_NAME,'form-type-textfield form-item-district form-disabled form-item form-group'.replace(' ', ".") )
    except Exception as e:
        а=0
        logging.info(f"У дела {dict_case_for_table['number_case']} не определено местоположение")

    try:
        dict_case_for_table['judicial_authority'] = get_value_from_element(driver, By.CLASS_NAME,'form-type-textfield form-item-court form-disabled form-item form-group'.replace(' ', ".") )
    except Exception as e:
        а=0
        logging.info(f"У дела {dict_case_for_table['number_case']} не определен судебный орган")

    try:
        dict_case_for_table['case_category'] = get_value_from_element(driver, By.CLASS_NAME,'form-type-textfield form-item-categoryGroup form-disabled form-item form-group'.replace(' ', ".") )
    except Exception as e:
        а=0
        logging.info(f"У дела {dict_case_for_table['number_case']} не определена категория")

    try:
        nature_of_statement = get_value_from_element(driver, By.CLASS_NAME,'form-type-textfield form-item-categoryGroup form-disabled form-item form-group'.replace(' ', ".") )
        if dict_case_for_table['case_category'] != '':
            dict_case_for_table['case_category'] = f"{dict_case_for_table['case_category']}, {nature_of_statement}"
        else:
            dict_case_for_table['case_category'] = nature_of_statement
    except Exception as e:
        а=0
        logging.info(f"У дела {dict_case_for_table['number_case']} не определен характер заявления")

    check_panels(driver, dict_case_for_table)


# Заполнение данных из табличных частей

# In[158]:



def check_panels(driver, dict_case_for_table: dict) -> None:
    # допустим, здесь уже открыт нужный кейс и загружены все панели
    # 1) находим все панели‑fieldset
    panels = driver.find_elements(
        By.CSS_SELECTOR,
        "fieldset.panel.panel-default.form-wrapper"
    )

    for panel in panels:
        # 2) пытаемся прочитать заголовок
        try:
            title_el = panel.find_element(By.CSS_SELECTOR, "legend .panel-title")
            title = title_el.text.strip().lower()

        except NoSuchElementException:
            # если вдруг нет заголовка — пропускаем
            continue

        # 3) в зависимости от заголовка обрабатываем нужную таблицу
        if title == "информация об оплате":
            try:
                # ищем таблицу внутри тела панели
                table = panel.find_element(By.CSS_SELECTOR, "div.panel-body table")
                # print(table.text)
                filling_amounts(table, dict_case_for_table)
            except NoSuchElementException as e:
                # если таблицы нет — вытягиваем данные полями формы
                get_payment_information_from_fields(panel, dict_case_for_table)
                logging.error("Ошибка в блоке 'информация об оплате': %s", e, exc_info=False)



        elif title == "стороны процесса":

            try:

                table = panel.find_element(By.CSS_SELECTOR, "div.panel-body div.table")
                # print('2', table.text)
                filling_parties_process(table, dict_case_for_table)
            except NoSuchElementException as e:
                # если таблицы сейчас нет, можно залогировать или пропустить
                msg = "стороны процесса" + str(e)
                logging.error("Ошибка в блоке 'стороны процесса': %s", e, exc_info=True)


# In[ ]:





# Заполнение данных из сторон процесса

# In[159]:




def filling_parties_process(defendants_table, dict_case_for_table:dict) -> None:

    list_defendants =[]
    parties_process = get_table_rows(defendants_table) # получение сторон процесса

    for process_side in parties_process[1:]: # проход по сторонам процесса исключив шапку

        side_name = get_column_value(process_side, 0, True)

        if side_name in ['ответчик', 'должник']:
            dict_defendant = create_dict_defendant()
            dict_defendant['defendant_fio'] = get_column_value(process_side, 3)
            dict_defendant['defendant_inn'] = get_column_value(process_side, 2)
            list_defendants.append(dict_defendant)

        elif side_name == 'представитель':
            dict_case_for_table['plaintiff_representative'] = get_column_value(process_side, 3)

    dict_case_for_table['list_defendants'] = list_defendants


# Заполнение сумм



def filling_amounts(amounts_table, dict_case_for_table: dict) -> None:
    # получаем только строки внутри <tbody>
    rows = amounts_table.find_elements(By.CSS_SELECTOR, 'tbody > tr')
    # пропускаем шапку, если по какой-то причине она там есть;
    # чаще всего в <thead>, но перестрахуемся
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) < 4:
            continue  # не та строка
        # чистим текст: убираем неразрывные пробелы и обычные
        claim = clean_text(cells[2].text)
        duties = clean_text(cells[3].text)
        dict_case_for_table['amount_of_claim'] = claim
        dict_case_for_table['amount_of_state_duties'] = duties


def clean_text(text: str) -> str:
    # удаляем все виды пробельных символов (включая &nbsp; → '\xa0')
    no_spaces = re.sub(r'\s+', '', text)
    return no_spaces  # здесь можно добавить re.sub(r'[^0-9,\.]', '', no_spaces) для чисто чисел


# Создание словаря для ответчика

# In[161]:


def create_dict_defendant() -> dict:
    cols = ('defendant_inn', 'defendant_fio')
    return dict.fromkeys(cols)


# Заполнить суммы из полей

# In[162]:


def get_payment_information_from_fields(panel, dict_case_for_table: dict) -> None:
    try:
        value_duty = panel.find_element(By.ID, 'edit-duty').get_attribute('value')
        dict_case_for_table['amount_of_state_duties'] = value_duty
    except NoSuchElementException:
        а=0
        # print(f"У дела {dict_case_for_table['number_case']} поле с государственной пошлиной не найдено")
    try:
        value_claim = panel.find_element(By.ID, 'edit-claim').get_attribute('value')
        dict_case_for_table['amount_of_claim'] = value_claim
    except NoSuchElementException:
        а=0
        # print(f"У дела {dict_case_for_table['number_case']} поле с суммой иска не найдено")


# In[163]:


# Переход на следующую старинцу.


# In[164]:


def go_to_next_page(driver):
    """
    Нажимает на кнопку «►» для перехода на следующую страницу.
    Ищет все ссылки внутри .list-pages и кликает ту, у которой текст ►.
    :param driver: Selenium WebDriver
    :return: True, если клик выполнен; False, если кнопка не найдена.
    """
    try:
        links = driver.find_elements(By.CSS_SELECTOR, '.list-pages a')
        for link in links:
            if link.text.strip() == '►':
                link.click()
                sleep(0.5)
                return True
        logger.error("Кнопка «►» не найдена среди ссылок.")
        return False
    except NoSuchElementException as e:
        logging.error("Контейнер .list-pages или ссылки внутри не найдены: %s", e, exc_info=True)
        return False


# Список записей



LIST_DICTS_CASES = [] # TODO !!!! # list(dict(create_dict_for_table()))





# # Проверка жив ли сайт



# длительное ожидание «возврата» сайта
SITE_WAIT_TOTAL_SECONDS = 10 * 60 * 60   # 10 часов
SITE_RETRY_INTERVAL_SECONDS = 2 * 60    # 2 минут

# чтобы Selenium не «вис» на долгих загрузках
PAGELOAD_TIMEOUT_SECONDS = 90            # например, 90 сек


def site_is_up(url: str, timeout: int = 15) -> bool:
    try:
        # HEAD быстрее, но не все конфигурации его обрабатывают; если что — используйте GET(stream=False)
        r = requests.head(url, timeout=timeout, allow_redirects=True)
        # считаем, что сайт «жив», если нет 5xx
        return r.status_code < 500
    except Exception:
        return False


"""
def wait_until_site_is_up(
    url: str,
    total_wait: int = SITE_WAIT_TOTAL_SECONDS,
    interval: int = SITE_RETRY_INTERVAL_SECONDS,
) -> bool:
    deadline = time.time() + total_wait
    attempt = 1
    while time.time() < deadline:
        if site_is_up(url):
            logger.info("Сайт доступен (попытка %d). Продолжаем.", attempt)
            return True
        remaining = max(0, int(deadline - time.time()))
        logger.error(
            "Сайт недоступен (попытка %d). Следующая проверка через %d мин. "
            "Осталось ждать ~%d мин.",
            attempt, interval // 60, remaining // 60
        )
        time.sleep(min(interval, remaining))
        attempt += 1
    logger.error("Сайт не восстановился за отведённое время (%d часов).", total_wait // 3600)
    return False
"""

'''


def wait_until_site_is_up(
    url: str,
    total_wait: int = SITE_WAIT_TOTAL_SECONDS,
    interval: int = SITE_RETRY_INTERVAL_SECONDS,
) -> bool:
    """
    Временный вариант. 
    """ 
    return True
'''

DOWN_SIGNATURES = [
    "bweb000065: http status 500",
    "jbweb000309: type",
    "jbweb000067: status report",
    "jbweb000145: the server encountered an internal error",
    "jboss web/7.2.2.final-redhat-1",
]
TRANSIENT_5XX = {502, 503, 504}

def _tcp_probe(url: str, timeout: float = 2.0) -> bool:
    """
    Быстрая проверка: DNS + TCP-подключение к host:port.
    Не делаем TLS-рукопожатие (чтобы не падать на нестандартных/корп. сертификатах).
    """
    parts = urlsplit(url)
    host = parts.hostname
    if not host:
        return False
    port = parts.port or (443 if parts.scheme == "https" else 80)
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception as e:
        logger.debug("tcp_probe failed: %r", e)
        return False


def _http_get(url: str, timeout: float = 5.0, verify_ssl: bool = False, max_bytes: int = 256 * 1024):
    """
    Аккуратный GET с коротким таймаутом. По умолчанию НЕ валидируем SSL,
    чтобы избежать ложных отказов из-за цепочек/прокси.
    Возвращает (status_or_None, body_text, err_or_None).
    """
    ctx = ssl.create_default_context()
    if not verify_ssl:
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

    req = Request(
        url,
        method="GET",
        headers={
            "User-Agent": "curl/8",   # многие сервера к этому лояльны
            "Accept": "*/*",
            "Connection": "close",
        },
    )
    try:
        with urlopen(req, timeout=timeout, context=ctx) as resp:
            code = resp.getcode()
            body = resp.read(max_bytes).decode("latin-1", errors="ignore")
            return code, body, None
    except HTTPError as e:
        # Это НЕ сетевой обрыв — сервер ответил валидным HTTP
        try:
            body = e.read(max_bytes).decode("latin-1", errors="ignore")
        except Exception:
            body = ""
        return e.code, body, None
    except Exception as e:
        return None, "", e


def site_is_up(url: str, timeout: int = 5, verify_ssl: bool = False) -> bool:
    """
    Критерии:
    - Если получили любой 2xx/3xx/4xx — ЖИВ (True).
    - 5xx:
        * 502/503/504 — ПАДАЕТ (False).
        * 500 — ПАДАЕТ только если совпала JBoss-сигнатура; иначе считаем ЖИВ (True).
    - Если HTTP не удался, но TCP-порт отвечает — считаем ЖИВ (True), чтобы не уходить в ложное ожидание.
    - Полный сетевой обрыв/нет DNS/порт закрыт — ПАДАЕТ (False).
    """
    tcp_ok = _tcp_probe(url, timeout=min(2.0, timeout))

    code, body, err = _http_get(url, timeout=timeout, verify_ssl=verify_ssl)

    if code is not None:
        if 200 <= code < 500:
            return True
        if code in TRANSIENT_5XX:
            return False
        if code >= 500:
            low = body.lower()
            if any(sig in low for sig in DOWN_SIGNATURES):
                return False
            return True
        # редкие коды <200 — трактуем как живой ответ
        return True

    # HTTP не удался (таймаут, SSL-ошибки и т.п.)
    if tcp_ok:
        return True
    return False


def wait_until_site_is_up(
    url: str,
    total_wait: int = 30 * 60,          # общее окно ожидания, сек
    interval: int = 5 * 60,             # интервал между «официальными» проверками, сек
    timeout_per_check: int = 5,         # таймаут на один HTTP-чек, сек
    consecutive_failures_threshold: int = 2,  # «дебаунс»: сколько подряд фейлов считать настоящим падением
) -> bool:
    """
    Дебаунсим ложные фейлы: пока нет N подряд неуспехов — не уходим в длинный sleep.
    Это резко снижает вероятность «ожидания», когда сайт фактически жив.
    """
    deadline = time.time() + total_wait
    attempt = 1
    fail_streak = 0

    while time.time() < deadline:
        if site_is_up(url, timeout=timeout_per_check, verify_ssl=False):
            logger.info("Сайт доступен (попытка %d). Продолжаем.", attempt)
            return True

        fail_streak += 1
        remaining = max(0, int(deadline - time.time()))

        if fail_streak < consecutive_failures_threshold:
            # Быстрый повтор — вдруг это единичный сбой сети/балансера.
            logger.warning(
                "Похоже, сбой (попытка %d, серия %d/%d). Перепроверяем через 1 сек. "
                "Осталось ждать ~%d мин.",
                attempt, fail_streak, consecutive_failures_threshold, max(0, remaining // 60)
            )
            time.sleep(1)
            attempt += 1
            continue

        # Настоящее «падение» — делаем паузу подлиннее
        logger.error(
            "Сайт недоступен (попытка %d, подряд %d). Следующая проверка через %d мин. "
            "Осталось ждать ~%d мин.",
            attempt, fail_streak, max(1, interval // 60), max(0, remaining // 60)
        )
        sleep_for = min(interval, remaining)
        if sleep_for <= 0:
            break
        time.sleep(sleep_for)
        attempt += 1

    logger.error(
        "Сайт не восстановился за отведённое время (%d часов).",
        max(1, total_wait // 3600)
    )
    return False


# # Сохранение локально для продолжение цикла.

# In[ ]:


import os
import pickle
import tempfile

def atomic_pickle_dump(obj, path):
    dirpath = os.path.dirname(path) or "."
    fd, tmppath = tempfile.mkstemp(dir=dirpath, prefix=".tmp.", suffix=".pkl")
    try:
        # 1) пишем во временный файл
        with os.fdopen(fd, "wb") as f:
            pickle.dump(obj, f, protocol=pickle.HIGHEST_PROTOCOL)
            f.flush()
            os.fsync(f.fileno())  # гарантируем запись на диск

        # 2) атомарно подменяем целевой файл
        os.replace(tmppath, path)  # atomic на POSIX и Windows (Py3.3+)

        # 3) (опционально, но полезно) fsync каталога — фиксируем сам rename
        try:
            dir_fd = os.open(dirpath, os.O_DIRECTORY)
            try:
                os.fsync(dir_fd)
            finally:
                os.close(dir_fd)
        except Exception:
            # не везде доступно, можно пропустить
            pass

    except Exception:
        # при ошибке уберём временный файл
        try:
            os.unlink(tmppath)
        except OSError:
            pass
        raise


# In[ ]:


SAVE_PATH = "parsing_state.pkl"

def save_globals():
    """Сохраняет глобальные переменные в файл"""
    state = {
        "current_page": current_page,
        "count_pages": count_pages,
        "count_page_now": count_page_now,
        "itaration_start_browser": itaration_start_browser,
        "current_case": current_case,
        "break_flag": break_flag,
        "LIST_DICTS_CASES": LIST_DICTS_CASES
    }
    atomic_pickle_dump(state, SAVE_PATH)

def load_globals():
    """Загружает глобальные переменные из файла"""
    global current_page, count_pages, count_page_now
    global itaration_start_browser, current_case, break_flag
    global LIST_DICTS_CASES, CASES_IS_LOAD, CASE_ERROR_COUNT, PROCESSED_CASE_NUMBERS

    with open(SAVE_PATH, "rb") as f:
        state = pickle.load(f)

    current_page = state["current_page"]
    count_pages = state["count_pages"]
    count_page_now = state["count_page_now"]
    itaration_start_browser = state["itaration_start_browser"]
    current_case = state["current_case"]
    break_flag = state["break_flag"]
    LIST_DICTS_CASES = state["LIST_DICTS_CASES"]
    CASE_ERROR_COUNT = {}
    CASES_IS_LOAD = [dict_case['number_case'] for dict_case in  state["LIST_DICTS_CASES"]]
    PROCESSED_CASE_NUMBERS = set()

def initialize_state(continue_parsing: bool):
    """Инициализирует глобальные переменные"""
    global current_page, count_pages, count_page_now
    global itaration_start_browser, current_case, break_flag
    global LIST_DICTS_CASES, CASES_IS_LOAD, CASE_ERROR_COUNT, PROCESSED_CASE_NUMBERS

    if continue_parsing and os.path.exists(SAVE_PATH):
        load_globals()
        print("Загружено сохраненное состояние.")
    else:
        current_page = 1
        count_pages = 0 + 2
        count_page_now = 1
        itaration_start_browser = 0
        current_case = 0
        break_flag = False
        LIST_DICTS_CASES = []
        CASE_ERROR_COUNT = {}
        CASES_IS_LOAD = []
        PROCESSED_CASE_NUMBERS = set()
        print("Начато новое состояние.")




# In[189]:


initialize_state(continue_parsing)


# ### Сам парсинг сайта

# In[ ]:


logger.info("start")
while current_page < count_pages:
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")

        options.add_argument("--disable-images")  # Major speed improvement
        options.add_argument("--disable-javascript-harmony-shipping")
        options.add_argument("--memory-pressure-off")
        with webdriver.Chrome(options=options) as driver:
            driver.set_page_load_timeout(PAGELOAD_TIMEOUT_SECONDS)
            if break_flag:
                break 

            if len(driver.window_handles) > 1:
                logger.info('Добавлена новая вкладка. Это условие остановки цикла.')
                break_flag = True
                break


            if not wait_until_site_is_up("https://office.sud.kz/index.xhtml"):
                # аккуратно выходим — сайт так и не поднялся
                raise SystemExit(1)


            logger.info("вход в личный кабинет.")
            login_to_site(driver) # вход в личный кабинет

            logger.info("вход в список дел.")
            go_to_workcab(driver) # вход в сипок дел

            count_pages = get_count_pages(driver) # получение общего количества страниц
            count_page_now = 1

            for count_page in range(current_page, count_pages+1): # проход по страницам
                if break_flag:
                    break 

                if len(driver.window_handles) > 1:
                    logger.info('Добавлена новая вкладка. Это условие остановки цикла.')
                    break_flag = True
                    break

                # Перелистывание на нужную страницу.
                diff_page = count_page - count_page_now
                if diff_page <= 3:
                    logger.info(f"change_page new = {diff_page}")
                    for _ in range(diff_page):
                        go_to_next_page(driver)

                else:
                    logger.info(f"change_page old {diff_page}")
                    go_to_page(driver,count_page) # переход к странице
                count_page_now = count_page

                # если переходим на следующую страницу, то обнуляем счетчик ошибок
                if current_page != count_page:
                    itaration_start_browser = 0
                    current_case = 0
                current_page = count_page

                count_cases = count_cases_on_page(driver) # получение количества дел на странице
                logger.info(f"count_cases={count_cases}")

                for count_case in range(current_case, count_cases): # проход по делам
                    if break_flag:
                        break 

                    if len(driver.window_handles) > 1:
                        logger.info('Добавлена новая вкладка. Это условие остановки цикла.')
                        break_flag = True
                        break

                    # при перезапуске страницы текущее дело пропустится  -- дело не пропускаем
                    current_case = count_case # + 1  -- дело не пропускаем


                    case_number = get_case_number_by_index(driver, count_case)
                    case_number_clean = case_number.replace('№','').replace(' ','')
                    logger.info(f"case_number={case_number}")
                    # case_number = get_case_number(driver) # номер дела # TODO??? Выдаёт пустоту.         
                    # Если сохранен тогда пропускаю. Если нет. То больше не смотрю есть ли сохраненные.
                    if case_number_clean in CASES_IS_LOAD:
                        # go_to_workcab(driver) # возращение к странице дел. Не надо. Мы смотрим есть ли дело в скаченных.
                        continue
                    else:
                        CASES_IS_LOAD = []

                    come_to_case(driver, count_case) # вход в дело

                    type_dynamics = type_of_case_dynamics(driver, case_number) # определяем тип окна статусов дела
                    if type_dynamics == '':
                        logger.info(f'У дела {case_number} иной тип статусов дела, либо статусы отсутствуют')
                        continue

                    dict_case_for_table = create_dict_for_table()
                    dict_case_for_table['number_case'] = case_number_clean


                    dynamics_case = get_dynamics_case(driver, type_dynamics) # получение список элементов записей 
                    recording_information_from_records(dynamics_case, dict_case_for_table, type_dynamics) # запись данных из динамики дела 

                    filling_in_additional_information(driver, dict_case_for_table) # запись данных из информации
                    go_to_workcab(driver) # возращение к странице дел

                    # если не было ни одно ответчика
                    if len(dict_case_for_table['list_defendants']) == 0:
                        dict_case_for_table['list_defendants'] = [create_dict_defendant()]

                    LIST_DICTS_CASES.append(dict_case_for_table)
                    save_globals()



    except Exception as e:
        logger.error('+'*30)
        logger.error(f'{datetime.datetime.today().strftime("%d.%m.%y %H:%M:%S")}: Ошибка на странице {current_page}.')
        logger.error(f"Ошибка: {e}", exc_info=True)
        if len(LIST_DICTS_CASES) > 0:
            logger.error(f'Последнее успешно загруженное дело: {LIST_DICTS_CASES[-1]["number_case"]}.')


        # подождать «возврата» сайта, но не бесконечно
        wait_until_site_is_up("https://office.sud.kz/index.xhtml")

        # Ошибка произошла в процессе обработки конкретного дела
        if 'case_number' in locals():
            key = case_number.strip()
            CASE_ERROR_COUNT[key] = CASE_ERROR_COUNT.get(key, 0) + 1
            logger.warning(f'Ошибка по делу {key}: попытка №{CASE_ERROR_COUNT[key]}')

            if CASE_ERROR_COUNT[key] >= 3:
                dict_case_for_table = create_dict_for_table()
                clean_case_number = key.replace('№','').replace(' ','')

                # проверка на дублирующийся номер
                if clean_case_number in PROCESSED_CASE_NUMBERS:
                    clean_case_number = 'error при записи'

                dict_case_for_table['number_case'] = clean_case_number
                # dict_case_for_table['defendant_fio'] = 'error'
                dict_case_for_table['plaintiff_representative'] = 'error'

                LIST_DICTS_CASES.append(dict_case_for_table)
                save_globals()

                PROCESSED_CASE_NUMBERS.add(clean_case_number)
                logger.warning(f'Добавлено ошибочное дело {clean_case_number}, переход к следующему.')
                current_case += 1
                continue


        itaration_start_browser += 1
        if itaration_start_browser > NUMBER_OF_BROWSER_REOPENINGS:
            logger.error(f'Прекращение выгрузки.')
            break 
        logger.error(f'Перезапуск выгрузки, начиная со страницы {current_page}.') #, пропуская {current_case} дело.') -- дело не пропускаем
        logger.error('-'*30)



# # Загрузка данных в Excel

# Имена колонок для Excel

# In[ ]:


DICT_ALIAS = {'number_case': '№ судебного дела', 
              'date_of_departure': 'Дата отправки искового заявления', 
              'date_of_rejection': 'Отклонено', 
              'cause_of_rejection': 'Причина отклонения заявления',
              'date_of_registration': 'Зарегистрировано',
              'name_judge': 'Судья',
              'date_of_return':'Вынесено определение о возврате искового заявления',
              'date_of_agreement':'Вынесено определение об утверждении соглашения об урегулировании спора',
              'date_of_simplification':'Вынесено определение о рассмотрении дела в порядке упрощенного производства',
              'date_of_first_instance':'Вынесено решение первой инстанции',
              #'index':'порядковый №',
              'plaintiff_representative': 'Представитель истца',
              'location': 'Область (столица, город республиканского значения)',
              'judicial_authority': 'Судебный орган',
              'case_category': 'Категория дела',
              'amount_of_claim': 'Сумма иска',
              'amount_of_state_duties': 'Сумма государственной пошлины',
              'defendant_inn': 'ИИН',
              'defendant_fio': 'ФИО',
              'date_court_order': 'Вынесено судебный приказ',
              'date_refusal_of_summary_proceedings': 'Определение об отмене решения в порядке упрощенного (письменного) производства',
             }


# загружаем данные в DataFrame

# In[ ]:


df = pd.json_normalize(LIST_DICTS_CASES, meta=LIST_COL)


# "Раскрываем" ответчиков

# In[ ]:


df = df.set_index(['plaintiff_representative', 	'number_case', 	'location', 	'judicial_authority', 	'case_category', 	'amount_of_claim',
                        'amount_of_state_duties', 	'date_of_departure', 	'date_of_rejection',  'cause_of_rejection',	'date_of_registration', 	'name_judge',
                        'date_of_return', 	'date_of_agreement', 	'date_of_simplification', 	'date_of_first_instance', 'date_court_order', 'date_refusal_of_summary_proceedings'])
df = df.apply(lambda x: x.explode()).reset_index()
df = df.join(pd.json_normalize(df.pop('list_defendants'))).drop_duplicates()
df.index = range(1,len(df)+1,1)


# # Получение максимальной даты.

# In[ ]:


columns_max_date = ['date_of_simplification',
         'date_of_first_instance',
         'date_court_order', 
         'date_refusal_of_summary_proceedings']


# Функция для извлечения максимальной даты
def extract_max_date(date_str):
    if not isinstance(date_str, str):  # Проверяем, что строка не пуста и это строка
        return None
    # Ищем все даты в формате ДД.ММ.ГГГГ
    date_matches = re.findall(r'\d{2}\.\d{2}\.\d{4}', date_str)
    if not date_matches:  # Если дат нет, возвращаем None
        return None
    # Конвертируем строки в объекты datetime
    dates = [datetime.datetime.strptime(date, '%d.%m.%Y') for date in date_matches]
    # Возвращаем максимальную дату
    return max(dates)


for col in columns_max_date:
    df[col] = df[col].apply(extract_max_date)


# Форматируем таблицу

# In[ ]:


#изменяем порядок
df = df[['defendant_inn',
         'defendant_fio',
         'plaintiff_representative',
         'number_case',
         'location',
         'judicial_authority',
         'case_category',
         'amount_of_claim',
         'amount_of_state_duties',
         'date_of_departure', 
         'date_of_rejection', 
         'cause_of_rejection',
         'date_of_registration',
         'name_judge',
         'date_of_return',
         'date_of_agreement',
         'date_of_simplification',
         'date_of_first_instance',
         'date_court_order', 
         'date_refusal_of_summary_proceedings']]

#приводим суммы к типу чисел
df.loc[df['amount_of_claim'] == '', 'amount_of_claim']= np.nan
df.loc[df['amount_of_state_duties'] == '', 'amount_of_state_duties']= np.nan
df['amount_of_claim'] = df['amount_of_claim'].apply(lambda x: str(x).replace(',','.')).fillna(0).astype(float)
df['amount_of_state_duties'] = df['amount_of_state_duties'].apply(lambda x: str(x).replace(',','.')).fillna(0).astype(float)

#переименовываем шапку
df.rename(columns=DICT_ALIAS, inplace=True)


# Загружаем данные в Excel с форматированием файла. По умолчанию путь - где лежит текущий файл, имя - дата и время выполнения файла

# In[ ]:


if PATH_TO_EXCELFILE == '':
    PATH_TO_EXCELFILE = './'
if FILE_NAME == '':
    FILE_NAME = f'{datetime.datetime.today().strftime("%d%m%y_%H%M%S")}'


# In[ ]:


with pd.ExcelWriter(f'{PATH_TO_EXCELFILE}{FILE_NAME}.xlsx', engine='xlsxwriter') as writer:#openpyxl

    df.style.set_properties(**{'text-align': 'center'}).to_excel(writer, index=True, startrow=0, header=True)

    workbook = writer.book
    table_format = workbook.add_format({'text_wrap': True, 'align':'center','text_wrap': True,
                                         'valign':'center', 'border_color': 'black', 'border':1})

    header_format = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA','text_wrap': True, 'align':'center',
                                         'valign':'vcenter', 'border_color': 'black', 'border':1})  

    worksheet = writer.sheets['Sheet1']

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num+1, value, header_format)
    worksheet.set_row(0, 60)

    (max_row, max_col) = df.shape
    worksheet.conditional_format(1, 0, max_row, max_col, {'type': 'no_errors', 'format': table_format})

    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 45)
    worksheet.set_column('D:D', 45)
    worksheet.set_column('E:E', 25)
    worksheet.set_column('F:F', 35)
    worksheet.set_column('G:G', 70)
    worksheet.set_column('H:H', 35)
    worksheet.set_column('I:I', 15)
    worksheet.set_column('J:J', 17)
    worksheet.set_column('K:K', 15)
    worksheet.set_column('L:L', 15)
    worksheet.set_column('M:M', 45)
    worksheet.set_column('N:N', 15)
    worksheet.set_column('O:O', 30)
    worksheet.set_column('P:P', 20)
    worksheet.set_column('Q:Q', 25)
    worksheet.set_column('R:R', 25)
    worksheet.set_column('S:S', 15)
    worksheet.set_column('T:T', 25)
    worksheet.set_column('U:U', 30)    


# # Создаю эксель с гугл таблицей

# In[ ]:


# df.to_pickle("df_start.pickle")
# df_google_t.to_pickle("google_table.pickle")
# df = pd.read_pickle("df_start.pickle")
# df_google_t = pd.read_pickle("google_table.pickle")


# In[ ]:


### Получение максимальных дат
columns_with_dates = ["Дата отправки искового заявления", 
                    "Зарегистрировано", 
                     "Вынесено определение о возврате искового заявления",
                    "Вынесено определение об утверждении соглашения об урегулировании спора",
                    "Отклонено"
                    ]

def split_date(s):
    if not isinstance(s, str):
        return s

    if s == "":
        return ""
    s = str(s)
    dates_str = s.split(", ")


    dates = [datetime.datetime.strptime(s_date, "%d.%m.%Y") for s_date in dates_str]
    max_date = max(dates)
    return max_date.date()



for col in columns_with_dates:
    df[col] = df[col].apply(split_date)
# columns_to_new_data_format = [
#     "Дата отправки искового заявления", 
#     "Отклонено",
#     "Зарегистрировано", 
#     "Вынесено определение о возврате искового заявления",
#     "Вынесено определение об утверждении соглашения об урегулировании спора",

#     "Вынесено определение о рассмотрении дела в порядке упрощенного производства", 
#     "Вынесено решение первой инстанции",
#     "Вынесен судебный приказ",
#     "Определение об отмене решения в порядке упрощенного производства"]


## Подтягиваю гугл таблицу
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
    # Получаем рабочий лист по названию
    worksheet = sh.worksheet(sheet_title)

    # Получаем все данные из листа
    data = worksheet.get_all_values()

    # Создаем DataFrame из полученных данных
    # Первая строка используется в качестве заголовков столбцов
    df = pd.DataFrame(data[1:], columns=data[0])

    return df


def get_google_df():
    gc: Client = gspread.service_account("./credentials kpi-api.json")
    sh: Spreadsheet = gc.open_by_url(SPREADSHEET_URL)

    show_available_worksheets(sh)    
    df = get_dataframe_from_sheet(sh, "ММ")
    return df

# df_google_t = get_google_df()
df_google_t.head(1)
df_google_t['Уникальный номер']


portfel_name = "Цессия" # "Продукт"
for col_need in [portfel_name, 'Уникальный номер', "ИИН"]:
    if col_need not in list(df_google_t):
        raise ValueError(f"Нет колонки {col_need}")

df_count_iin = df_google_t.groupby(["ИИН"], as_index=False).agg({'Уникальный номер':"count"}).rename(columns={'Уникальный номер':"количество_записей_по_ИИН"})
df_google_t = pd.merge(df_google_t, df_count_iin, on=["ИИН"], how="left")
df = pd.merge(df, df_count_iin, on=["ИИН"], how="left")
# df.to_pickle("df.pickle")
df_google_t["количество_записей_по_ИИН"] = df_google_t["количество_записей_по_ИИН"].astype(np.int32)
## Присоедение для кол-ва ИИН == 1 
df_google_temp = df_google_t.loc[df_google_t["количество_записей_по_ИИН"]==1, 
                                 ['Уникальный номер', "ИИН"]]
df = pd.merge(df, df_google_temp, on=["ИИН"], how="left")
# Удаляю дублирующие строки для ИИН у которых 1 запись в гугл таблице.

mask = df['количество_записей_по_ИИН'] == 1          # строки, где нужно убрать дубликаты
dups = df.duplicated(subset=["ИИН"], keep='first')  # True для всех «задвоенных» строк
                                               # оставит первый экземпляр

# оставляем всё, кроме тех строк, где И 1. mask, И 2. dups = True
df_no_duplicate = df[~(mask & dups)].copy()

df_no_duplicate.shape, df.shape
df_no_duplicate["123"] = 1
df_temp = df_no_duplicate.groupby(["Уникальный номер"]).agg({"123":"count"})
df_temp.loc[df_temp["123"] != 1]
## Присоедение для кол-ва ИИН >= 2 
### Для 7517
df_no_duplicate['Уникальный номер'].isna().sum()
df_no_duplicate['№ судебного дела_4symbol'] = df_no_duplicate['№ судебного дела'].apply(lambda x: x[:4])

def split_name_product(s):
    first_words = re.split(r'[\s.-]', s, maxsplit=1)[0]   # ⟵ первый «кусочек»
    return first_words


df_google_t["Продукт_short"] = df_google_t[portfel_name].apply(split_name_product)
name_uid = "Уникальный номер2"
name_portfolio_with_7517 = ['ММ', 'Solva', 'LIME', 'Bereke']

df_google_t['тип_продукта'] = 0
df_google_t.loc[df_google_t["Продукт_short"].isin(name_portfolio_with_7517), ['тип_продукта']] = 7517

df_count_iin_and_type = (df_google_t.groupby(["ИИН", 'тип_продукта'], 
                                             as_index=False)
                                     .agg({"Уникальный номер":"count"})
                                     .rename(columns={"Уникальный номер":"count_iin_and_one_type"})
                        )

if "count_iin_and_one_type" not in list(df_google_t):
    df_google_t = pd.merge(df_google_t, df_count_iin_and_type[["ИИН", 'тип_продукта', "count_iin_and_one_type"]], on=["ИИН", 'тип_продукта'], how='left')


df_google_temp_7517 = df_google_t.loc[(df_google_t["количество_записей_по_ИИН"]>=2)&
                                (df_google_t['count_iin_and_one_type']== 1)&
                                 (df_google_t["тип_продукта"] == 7517), 
                                 ['Уникальный номер', "ИИН", "Продукт_short"]].copy()


df_google_temp_7517['№ судебного дела_4symbol'] = "7517"

df_google_temp_7517 = df_google_temp_7517.rename(columns={"Уникальный номер":name_uid})
print(df_no_duplicate.shape)
df_no_duplicate = pd.merge(df_no_duplicate, df_google_temp_7517, 
              left_on=["ИИН", "№ судебного дела_4symbol"], 
              right_on=["ИИН", "№ судебного дела_4symbol"], how='left')

print(df_no_duplicate.shape)
print("Пропусков до:", df_no_duplicate.loc[:, 'Уникальный номер'].notna().sum())

mask = df_no_duplicate['Уникальный номер2'].notna()
df_no_duplicate.loc[mask, 'Уникальный номер'] = df_no_duplicate.loc[mask, 'Уникальный номер2']

print("Пропусков после:", df_no_duplicate.loc[:, 'Уникальный номер'].notna().sum())
### Для не 7517 с 2-мя ИИН
name_uid = "Уникальный номер2.2"

df_google_temp_0 = df_google_t.loc[(df_google_t["количество_записей_по_ИИН"]>=2)&
                                (df_google_t['count_iin_and_one_type']== 1)&
                                 (df_google_t["тип_продукта"] == 0), 
                                 ['Уникальный номер', "ИИН", "Продукт_short"]].copy()


df_google_temp_0["№ судебного дела_4symbol_clean"] = "0"
df_google_temp_0 = df_google_temp_0.rename(columns={"Уникальный номер":name_uid})
df_no_duplicate["№ судебного дела_4symbol_clean"] = df_no_duplicate["№ судебного дела_4symbol"].apply(lambda x: x if x == "7517" else "0")
print(df_no_duplicate.shape)
df_no_duplicate = pd.merge(df_no_duplicate, df_google_temp_0[[name_uid, "ИИН", "№ судебного дела_4symbol_clean"]], 
              left_on=["ИИН", "№ судебного дела_4symbol_clean"], 
              right_on=["ИИН", "№ судебного дела_4symbol_clean"], how='left')

print(df_no_duplicate.shape)
print("Заполнено до:", df_no_duplicate.loc[:, 'Уникальный номер'].notna().sum())

mask = df_no_duplicate[name_uid].notna()
df_no_duplicate.loc[mask, 'Уникальный номер'] = df_no_duplicate.loc[mask, name_uid]

print("Заполнено после:", df_no_duplicate.loc[:, 'Уникальный номер'].notna().sum())
# Оставлю или не дубли по уникальному, или первый элемент в дублях.

mask = df_no_duplicate['Уникальный номер'].notna()        # строки, где нужно убрать дубликаты
dups = df_no_duplicate.duplicated(subset=['Уникальный номер'], keep='first')  # True для всех «задвоенных» строк
# оставит первый экземпляр

# оставляем всё, кроме тех строк, где И 1. mask, И 2. dups = True
df_no_duplicate2 = df_no_duplicate[~(mask & dups)].copy()



### случайное заполнение уникального - Для оставшихся пропусков с 2мя ИИН 
df_google_join = df_google_t[['Уникальный номер', 'ИИН']].copy()

mask = ~(df_google_join["Уникальный номер"].isin(df_no_duplicate2['Уникальный номер']))

df_google_join = df_google_join.loc[mask]
df_google_join = df_google_join.rename(columns={"Уникальный номер":"Уникальный номер_random"})

df_google_join = df_google_join.drop_duplicates(["ИИН"], keep="last")
df_no_duplicate2["Возможно ошибочный уникальный"] = ""

mask = df_no_duplicate2["Уникальный номер"].isna()

df_no_duplicate2.loc[mask, "Возможно ошибочный уникальный"] = '!!!'
df_no_duplicate3 = pd.merge(df_no_duplicate2, df_google_join, on="ИИН", how='left')
df_no_duplicate2.shape, df_no_duplicate3.shape
mask = df_no_duplicate3['Уникальный номер'].isna()

df_no_duplicate3.loc[mask, 'Уникальный номер'] = df_no_duplicate3.loc[mask, 'Уникальный номер_random']
df_no_duplicate3['Уникальный номер'].isna().sum()
# Итоговый файл соеденённый с гугл таблицей
df_g = df_no_duplicate3.copy()
null_columns = ["Дата подачи заявления на выписку ИЛ", 
"Дата получения ИЛ", 
"Дата передачи ИЛ ЧСИ", 
"Комментарии по суду", 
"Дата подачи заявления на выдачу СП", 
"Дата получения СП", 
"Дата передачи СП ЧСИ", 
"Комментарий", 

"Ответственный", "Дата создания", "Название процесса"

]
for col in null_columns:
    df_g[col] = np.nan
df_g = df_g.rename(columns={'Уникальный номер':"Уникальный номер сделки", 
                            "№ судебного дела":"Номер судебного дела",
                            "Вынесено судебный приказ":"Вынесен судебный приказ",
                            "Определение об отмене решения в порядке упрощенного (письменного) производства":"Определение об отмене решения в порядке упрощенного производства"
                           })
df_g["Тип процесса"] = "1. Упрощённое производство"
df_g["Статус процесса"] = "Рассмотрение дела"
df_g.head(2)
df_g = df_g[["Уникальный номер сделки",

    "Дата подачи заявления на выписку ИЛ", 
    "Дата получения ИЛ", 
    "Дата передачи ИЛ ЧСИ", 
    "Комментарии по суду", 
    "Дата подачи заявления на выдачу СП", 
    "Дата получения СП", 
    "Дата передачи СП ЧСИ", 
    "Комментарий", 

    "Представитель истца",
      "Номер судебного дела",
      "Судебный орган",
      "Категория дела",
      "Сумма иска",
      "Сумма государственной пошлины",
      "Дата отправки искового заявления", 
      "Отклонено",
      "Причина отклонения заявления",
      "Зарегистрировано",
      "Судья",
      "Вынесено определение о возврате искового заявления",
      "Вынесено определение об утверждении соглашения об урегулировании спора",
      "Вынесено определение о рассмотрении дела в порядке упрощенного производства",
      "Вынесено решение первой инстанции",
      "Вынесен судебный приказ",
      "Определение об отмене решения в порядке упрощенного производства",


      "Ответственный", 
      "Дата создания", 
      "Название процесса",

      "Тип процесса", 
      "Статус процесса",
      "ИИН",
      "Возможно ошибочный уникальный"

    ]]
# df_g["123"] = 1
# df_temp = df_g.groupby(["Уникальный номер сделки"]).agg({"123":"count"})
# df_temp.loc[df_temp["123"] != 1]


# In[ ]:


# Create a Pandas Excel writer using XlsxWriter as the engine.
# Also set the default datetime and date formats.
writer = pd.ExcelWriter(
    Path(PATH_TO_EXCELFILE) / FILE_NAME + "Суд с гугл таблицей.xlsx",
    engine="xlsxwriter",
    datetime_format="dd.mm.yyyy",
    date_format="dd.mm.yyyy",
)

# Convert the dataframe to an XlsxWriter Excel object.
df_g.to_excel(writer, sheet_name="Sheet1", index=False)
writer.close()


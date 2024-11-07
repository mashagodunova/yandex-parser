import requests
import os
import openpyxl
from io import BytesIO
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
token = '5643122315:AAHTc3MOGgFiwXznLctdNPVRZ6QDO_iDSSs'


def create_driver():
    options = Options()
    options.add_argument(
        f"user-data-dir=Profile_14")
    options.add_argument('--user-agent="Mozilla/5.0 (Windows Phone 10.0; Android 4.2.1; Microsoft; Lumia 640 XL LTE) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Mobile Safari/537.36 Edge/12.10166"')
    #options.add_argument("--headless")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--v=99")
    options.add_argument("--no-sandbox")
    driver = uc.Chrome(options=options)
    return driver

def preprocessing(spis):
    for_file = []
    for el in spis:
        c_d={}
        nel = el.split('\n')
        c_d['Автор'] = nel[0]
        c_d['Дата'] = nel[2]
        if 'Товар продавца' in el:
            c_d['Продавец'] = nel[3]
        else:
            c_d['Продавец'] = 'None'
        if 'Достоинства' in el and 'Недостатки' in el and 'Товар продавца' in el:
            c_d['Достоинства'] = nel[nel.index('Достоинства')+1]
            c_d['Недостатки'] = nel[nel.index('Недостатки')+1]
        elif 'Достоинства' in el and 'Недостатки' in el and 'Товар продавца' not in el:
            c_d['Достоинства'] = nel[nel.index('Достоинства')+1]
            c_d['Недостатки'] = nel[nel.index('Недостатки')+1]
        elif 'Достоинства' not in el and 'Недостатки' not in el:
            c_d['Достоинства'] = 'None'
            c_d['Недостатки'] = 'None'
        elif 'Достоинства' not in el and 'Недостатки' in el and 'Товар продавца' in el:
            c_d['Достоинства'] = 'None'
            c_d['Недостатки'] = nel[nel.index('Недостатки')+1]
        elif 'Достоинства' in el and 'Недостатки' not in el and 'Товар продавца' in el:
            c_d['Достоинства'] = nel[nel.index('Достоинства')+1]
            c_d['Недостатки'] = 'None'
        elif 'Достоинства' not in el and 'Недостатки' in el and 'Товар продавца' not in el:
            c_d['Достоинства'] = 'None'
            c_d['Недостатки'] = nel[nel.index('Недостатки')+1]
        elif 'Достоинства' in el and 'Недостатки' not in el and 'Товар продавца' not in el:
            c_d['Достоинства'] = nel[nel.index('Достоинства')+1]
            c_d['Недостатки'] = 'None'
        if 'Комментарий' in el:
            c_d['Комментарий'] = nel[nel.index('Комментарий')+1]
        elif 'Комментарий' not in el:
            c_d['Комментарий'] = 'None'
        c_d['Лайки'] = nel[-3]
        for_file.append(c_d)
    return for_file

def parser_yandex_market(link):
    driver = create_driver()
    driver.get(link)
    time.sleep(10)
    #EQlfk COmSo
    reviews_link = driver.find_element(By.XPATH,"//a[@class='_3RJHd _24IFf']").get_attribute('href')
    reviews_link=reviews_link+'&page=1'
    driver.quit()
    time.sleep(2)
    try:
        link = reviews_link
        driver = create_driver()
        time.sleep(4)
        driver.get(link)
        time.sleep(10)
        # button = driver.find_element(By.XPATH,"//button[@class='_3CCE- _1Mcpo _2jQ3e _3ozc8 _33GKh']")
        # button.click()
        driver.execute_script("window.scrollTo(0, 50000)")
        final = []
        for el in driver.find_elements(By.XPATH,"//li[@class='_3sEW7']"):
            time.sleep(3)
            print(el.text)
            time.sleep(2)
            final.append(el.text)
        driver.quit()
        return preprocessing(final)
    except Exception as e:
        print(f'Ошибка парсинга {e}')
        driver.quit()
        return 'error'


def start_parser(chat_id, link):
    quetstion_func = parser_yandex_market(link)
    if quetstion_func == 'error':
        text = 'Ошибка при парсинге, повторите попытку'
        requests.post(
            f'https://api.telegram.org/bot{token}/sendMessage?chat_id={chat_id}&text={text}')
    else:
        df = pd.DataFrame(quetstion_func)
        # f = open("result_file.xlsx", 'a')
        df.to_excel('result_file.xlsx')
        f = open("result_file.xlsx", 'rb')
        status = requests.post(
            f'https://api.telegram.org/bot{token}/sendDocument?chat_id={chat_id}', files={'document': f})
        print(status)
        return status
        # os.remove(f)
#start_parser('741541899', 'https://market.yandex.ru/product--nabor-iz-10-par-noskov-moscowsocksclub-m20-miks-iskusstvo-razmer-27-41-43/1439308463?glfilter=14474426%3AMjcgKDQxLTQzKQ_101447518663&glfilter=14871214%3A15098577_101447518663&cpc=zHewvSNgIjPtJC9iD60JL3m43OcnhUqQ-g6C4kDb05rWnvIKBybbR6unOSuLlBsL9g15fnalFtreycesyu-TgSxkTsBYWCoNIxS5QvFMzLDWg2VG_56D7oYSPcsjO6rlKiAnV7kSCqmriKKOOtdM0Vd8TVPSkApivEnqem6eGihq0VwesfusPWNxg1F5Tq3i&sku=101447518663&offerid=7JFFcigWlvMzkapCVGazig&cpa=1')


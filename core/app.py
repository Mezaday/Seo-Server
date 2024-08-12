import time
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
import re
import requests
import random


class Application:

    def create_app(self) -> FastAPI:
        instance = FastAPI(
            title='SEO Server',
            version='0.0.1a'

        )
        return instance


app: FastAPI = Application().create_app()
templates = Jinja2Templates(directory='templates')
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/home", response_class = HTMLResponse)
def root(request: Request):
    df = pd.read_excel('data\данные.xlsx', sheet_name='main')
    artic = df['Артикул'].unique()
    return templates.TemplateResponse(
            name='home.html', request=request, context={"artics": artic}
    )


@app.post("/home", response_class = HTMLResponse)
def result(request: Request,
            seo: list = Form(),
            select: list = Form()):
    df_main = pd.read_excel('data\данные.xlsx', sheet_name='main', dtype={'Артикул': int, 'Брэнд': str})
    df_main = df_main.query(f'Артикул == {select[0]}')
    df_main = df_main.dropna().reset_index(drop=True)
    df = pd.read_excel('data\данные.xlsx', sheet_name=select[0], dtype={'Запрос': str, 'Частота': str})
    df = df.dropna().reset_index(drop=True)
    s = ''
    for el in seo:
        s += str(el)
        s += "."
        s += " "
    dict = []
    for words in df['Запрос'].to_list():
        pattern = re.escape(words)
        if re.search(pattern, s, flags=re.I):
            num = {'1': words, '2': '+', '3': '', '4': '', '5': df.query(f'Запрос == "{words}"').Частота.iloc[0]}
            dict.append(num)
        else:
            wordes = words.split(' ')
            for word in wordes:
                pattern = re.escape(word)
                if re.search(pattern, s, flags=re.I):
                    a = True
                else:
                    a = False
                    break
            if a == False:
                num = {'1': words, '2': '', '3': '', '4': '-', '5': df.query(f'Запрос == "{words}"').Частота.iloc[0]}
                dict.append(num)
            else:
                num = {'1': words, '2': '', '3': '±', '4': '', '5': df.query(f'Запрос == "{words}"').Частота.iloc[0]}
                dict.append(num)
    i = 0
    t = 0
    check = []
    while i < df.shape[0]:
        # print(f'i = {i}')
        brand = df_main.Брэнд[0]
        key_word = df.Запрос[i]
        ID = df_main.Артикул[0]
        sess = requests.Session()
        proxy = [
            {
                'http': 'proxy1',
                'https': 'proxy1',

            },
            {
                'http': 'proxy2',
                'https': 'proxy2',
            },
            {
                'http': 'proxy3',
                'https': 'proxy3',
            },
            {
                'http': 'proxy4',
                'https': 'proxy4',
            }

        ]
        proxies = random.choice(proxy)
        url = (f"https://search.wb.ru/exactmatch/ru/common/v4/search?"
               f"appType=1&curr=rub&dest=-1257786"
               f"&brand={'%20'.join(brand.split())}&page=1"
               f"&query={'%20'.join(key_word.split())}&resultset=catalog"
               f"&sort=popular&spp=29&suppressSpellcheck=false")
        response = sess.get(url, headers={'Accept': "*/*",
                                          'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                                                                             " AppleWebKit/537.36 (KHTML, like Gecko)"
                                                                             " Chrome/120.0.0.0 YaBrowser/24.1.0.0 Safari/537.36"},
                                          proxies=proxies)
        if response.status_code != 200:
            print(f'Код ошибки {response.status_code}')
            break
        response_JSON = response.json()
        products_on_page = []
        metadata = response_JSON['metadata']
        if response_JSON['data']['products'] == []:
            i = i + 1
            # print(f'i = {i}')
            for data in dict:
                if data.get('1') == key_word:
                    num = data
                    check.append(num)
                    time.sleep(3)
        elif metadata['name'] != key_word:
            print(response_JSON)
            name = metadata['name']
            search = key_word.split(' ')
            for key_words in search:
                pattern = re.escape(key_words)
                if re.search(pattern, s, flags=re.I):
                    a = True
                else:
                    a = False
                    break
            if a == True:
                for item in response_JSON['data']['products']:
                    products_on_page.append({
                        'Артикул': item['id']
                    })
                    df_check = pd.DataFrame.from_dict(products_on_page)
                q = ID in df_check['Артикул'].tolist()
                # print(q)
            else:
                q = False
        else:
            for item in response_JSON['data']['products']:
                products_on_page.append({
                    'Артикул': item['id']
                })
                df_check = pd.DataFrame.from_dict(products_on_page)
            q = ID in df_check['Артикул'].tolist()
            # print(q)
        if q == False:
            t = t + 1
            # print(f't = {t}')
            if t != 3:
                time.sleep(2)
                continue
            else:
                i = i + 1
                t = 0
                # print(response_JSON)
                for data in dict:
                    if data.get('1') == key_word:
                        num = data
                        check.append(num)
                        time.sleep(2)
        else:
            i = i + 1
            t = 0
            time.sleep(2)

    exсel = pd.DataFrame.from_dict(check)

    df_final = pd.DataFrame()
    df_final['Запрос'] = exсel['1']
    df_final['+'] = exсel['2']
    df_final['±'] = exсel['3']
    df_final['-'] = exсel['4']
    df_final['Частота'] = exсel['5']

    writer = pd.ExcelWriter(
        "data\Не прошедшие проверку.xlsx",
        engine="xlsxwriter",
        datetime_format="mm-dd-yyyy",
        date_format="mm-dd-yyyy",
    )
    df_final.to_excel(writer, sheet_name="Sheet1")
    worksheet = writer.sheets["Sheet1"]

    (max_row, max_col) = df_final.shape

    worksheet.set_column(1, max_col, 14)

    writer.close()
    return templates.TemplateResponse(
        name='result.html', request=request, context={"dict": dict, "seo": seo, "select": select, "check": check}
    )
from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import xlwt
import time
import random
import socks
import socket

# работаем под запущенным Tor-браузером
socks.set_default_proxy(socks.SOCKS5, "localhost", 9150)
socket.socket = socks.socksocket

useragent = UserAgent()

names = []

for i in range(1, 222): # 221 страниц
    url = 'https://www.etsy.com/shop/RStudioDesign/reviews?ref=pagination&page=' + str(i)
    response = requests.get(url, headers={'User-Agent': useragent.random})  # подменяем агент
    print(response.status_code, i)  # статус сервера (должен быть 200)
    # если статус не 200, делаем паузу 60-80 сек и снова делаем запрос:
    if response.status_code != 200:
        time.sleep(random.randrange(60, 80, 1))
        url = 'https://www.etsy.com/shop/RStudioDesign/reviews?ref=pagination&page=' + str(i)
        response = requests.get(url, headers={'User-Agent': useragent.random})
        print('  ', response.status_code, i)

    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    names_soup = soup.find_all('div', class_='flag-body hide-xs hide-sm')
    for name in names_soup:
        name = name.find('p').text
        names.append(name)

# Далее записываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('RStudioDesign')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 0
for i in range(len(names)):
    sheet.write(m, 0, names[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 20000  # Ширина колонки

sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('Shop1_RStudioDesign.xls')  # Сохраняем в файл

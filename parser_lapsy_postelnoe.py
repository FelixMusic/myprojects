from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import xlwt
import time
import random
import socks
import socket

socks.set_default_proxy(socks.SOCKS5, "localhost", 9150)
socket.socket = socks.socksocket

useragent = UserAgent()

links = []
artikul = []
name = []
price = []
price_discount = []
brand = []

def checkIP():
    ip = requests.get('http://checkip.dyndns.org').content
    soup = BeautifulSoup(ip, 'html.parser')
    print(soup.find('body').text)

for i in range(1, 51): # 50 страниц
    url = 'https://lapsi.ru/detskaya_komnata/tekstil/?PAGEN_5=' + str(i)
    response = requests.get(url, headers={'User-Agent': useragent.random})  # подменяем агент
    print(response.status_code, i)  # статус сервера (должен быть 200)

    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    links_soup = soup.find_all('div', class_='product-card__title')
    for link in links_soup:
        link = link.find('a').get('href')
        links.append('https://lapsi.ru' + link)
    time.sleep(1)
    checkIP()

time.sleep(10)
j = 1
for link in links:
    url = link
    response = requests.get(url, headers={'User-Agent': useragent.random})
    time.sleep(random.randrange(1, 3, 1))  # задержка времени от 3 до 7 сек
    print(response.status_code, j)
    html = response.content
    soup = BeautifulSoup(html, 'html.parser')
    try:
        page_soup = soup.find('div', class_='info clear').find('h1')
        name_page = page_soup.text
        name.append(name_page)
    except:
        time.sleep(11)
        response = requests.get(url, headers={'User-Agent': useragent.random})
        html = response.content
        soup = BeautifulSoup(html, 'html.parser')
        page_soup = soup.find('div', class_='info clear').find('h1')
        name_page = page_soup.text
        name.append(name_page)

    artikul_soup = soup.find('div', class_='product-property__value')
    artikul_item = artikul_soup.text
    artikul.append(artikul_item)

    price_soup = soup.find('div', class_='price')
    price_item = price_soup.text.replace('\xa0', '').replace(' ', '')
    price.append(int(price_item))

    try:
        brand_soup = soup.find('p', class_='product__description-only-shop').find('a')
        brand_item = brand_soup.text
        brand.append(brand_item)
    except:
        brand.append('-')

    try:
        price_discount_soup = soup.find('div', class_='old-price')
        price_discount_item = price_discount_soup.text.replace('\xa0', '').replace(' ', '')
        price_discount.append(int(price_discount_item))
    except:
        price_discount.append(0)
    j = j + 1

# Далее записываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('лапси постельное белье')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 0
for i in range(len(links)):
    sheet.write(m, 0, name[i], font)
    sheet.write(m, 1, artikul[i], font)
    sheet.write(m, 2, brand[i], font)
    sheet.write(m, 3, price[i], font)
    sheet.write(m, 4, price_discount[i], font)
    sheet.write(m, 5, links[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 22000  # Ширина колонки
sheet.col(1).width = 3000
sheet.col(2).width = 5000
sheet.col(3).width = 3000
sheet.col(4).width = 3000
sheet.col(5).width = 26000

sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('lapsy_postelnoe.xls')  # Сохраняем в файл

from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import xlwt
import time
import datetime
import lxml
import random

useragent = UserAgent()

links = []
artikul = []
name = []
price = []
price_discount = []
brand = []

for i in range(1, 20): # 19 страниц
    url = 'https://www.wildberries.ru/catalog/0/search.aspx?subject=' \
          '883&search=%D0%BF%D0%BE%D1%81%D1%82%D0%B5%D0%BB%D1%8C%D0%BD%' \
          'D0%BE%D0%B5%20%D0%B1%D0%B5%D0%BB%D1%8C%D0%B5&page=' + str(i)
    response = requests.get(url, headers={'User-Agent': useragent.random})  # подменяем агент
    print(response.status_code, i)  # статус сервера (должен быть 200)

    html = response.content
    soup = BeautifulSoup(html, 'lxml')

    links_soup = soup.find_all('div', class_='dtList i-dtList j-card-item')
    for link in links_soup:
        link = link.find('a').get('href')
        links.append(link)
    time.sleep(3)

time.sleep(5)

j = 1
for link in links:
    url = link
    response = requests.get(url, headers={'User-Agent': useragent.random})
    time.sleep(1)
    print(response.status_code, j)
    html = response.content
    soup = BeautifulSoup(html, 'lxml')
    page_soup = soup.find('div', class_='card-row').find('span', class_='name')
    name_page = page_soup.text
    name.append(name_page)

    try:
        brand_soup = soup.find('div', class_='card-row').find('span', class_='brand')
        brand_item = brand_soup.text
        brand.append(brand_item)
    except:
        brand.append('-')

    try:
        artikul_soup = soup.find('div', class_='article').find('span', class_='j-article')
        artikul_item = artikul_soup.text
        artikul.append(artikul_item)
    except:
        artikul.append('-')

    try:
        price_soup = soup.find('div', class_='final-price-block').find('span', class_='final-cost')
        price_item = price_soup.text.strip().replace('\xa0', '').replace('₽', '')
        price.append(price_item)
    except:
        price.append('-')

    try:
        price_discount_soup = soup.find('span', class_='old-price').find('del', class_='c-text-base')
        price_discount_item = price_discount_soup.text.replace('\xa0', '').replace('₽', '')
        price_discount.append(price_discount_item)
    except:
        price_discount.append(0)
    j = j + 1

# имя файла xls с текущей датой
date_now = datetime.datetime.now()
date_now_list = list(map(str, str(date_now).split()))
now_list = list(map(str, str(date_now_list[0]).split('-')))
datefld = str(now_list[2]) + '.' + str(now_list[1]) + '.' + str(now_list[0])
xls_file_name = 'WB_postelnoe_' + datefld + '.xls'

# Далее зиписываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('WB постельное')
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
book.save(xls_file_name)  # Сохраняем в файл

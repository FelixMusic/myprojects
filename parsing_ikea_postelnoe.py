from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import xlwt
import time
import lxml
import random

useragent = UserAgent()

links = []
artikul = []
name = []
pip_name = []
price = []
product_unit = []

for i in range(1, 3): # 2 страниц
    url = 'https://www.ikea.com/ru/ru/cat/tekstil-dlya-malyshey-18690/?page=' + str(i)
    response = requests.get(url, headers={'User-Agent': useragent.random})  # подменяем агент
    print(response.status_code, i)  # статус сервера (должен быть 200)

    html = response.content
    soup = BeautifulSoup(html, 'lxml')

    links_soup = soup.find_all('div', class_='product-compact__spacer')
    for link in links_soup:
        link = link.find('a').get('href')
        links.append(link)
    time.sleep(3)

j = 1
for link in links:
    url = link
    response = requests.get(url, headers={'User-Agent': useragent.random})
    time.sleep(1)
    # time.sleep(random.randrange(5, 8, 1))  # задержка времени от 3 до 7 сек
    print(response.status_code, j)
    html = response.content
    soup = BeautifulSoup(html, 'lxml')
    page_soup = soup.find('div', class_='product-pip__product-heading-container').find('span', class_='normal-font range__text-rtl')
    name_page = page_soup.text.strip()
    name.append(name_page)

    pip_name_soup = soup.find('div', class_='product-pip__product-heading-container').find('span', class_='product-pip__name')
    pip_name_current = pip_name_soup.text
    pip_name.append(pip_name_current)

    try:
        artikul_soup = soup.find('span', class_='range-product-identifier__number')
        artikul_item = artikul_soup.text
        artikul.append(artikul_item)
    except:
        artikul.append('-')

    try:
        price_soup = soup.find('span', class_='product-pip__price__value')
        price_item = price_soup.text.replace(' ₽', '').replace(' ', '')
        price.append(price_item)
    except:
        price.append('-')

    try:
        unit_soup = soup.find('span', class_='product-pip__price__unit')
        unit_item = unit_soup.text.replace('/', '')
        product_unit.append(unit_item)
    except:
        product_unit.append('-')
    j = j + 1

# Далее зиписываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('ikea постельное белье')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 0
for i in range(len(links)):
    sheet.write(m, 0, name[i], font)
    sheet.write(m, 1, artikul[i], font)
    sheet.write(m, 2, pip_name[i], font)
    sheet.write(m, 3, price[i], font)
    sheet.write(m, 4, product_unit[i], font)
    sheet.write(m, 5, links[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 22000  # Ширина колонки
sheet.col(1).width = 3000
sheet.col(2).width = 5000
sheet.col(3).width = 3000
sheet.col(4).width = 2000
sheet.col(5).width = 26000

sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('ikea_postelnoe.xls')  # Сохраняем в файл

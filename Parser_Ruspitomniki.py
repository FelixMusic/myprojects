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

first_links = []
second_links =[]
third_links = []
pictures = []

url = 'https://www.ruspitomniki.ru/catalog/index.html'
response = requests.get(url, headers={'User-Agent': useragent.random})  # подменяем агент
print(response.status_code)  # статус сервера (должен быть 200)

html = response.content
soup = BeautifulSoup(html, 'html.parser')

links_soup = soup.find_all('span', class_='col-pad')
for link in links_soup:
    link = link.find('a').get('href')
    first_links.append('https://www.ruspitomniki.ru' + link)


# print(first_links)

for first_link in first_links[0:2]:
    url = first_link
    response = requests.get(url, headers={'User-Agent': useragent.random})
    print(response.status_code)
    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    links_soup = soup.find_all('span', class_='mehrPad')
    for link in links_soup:
        link = link.find('a').get('href')
        second_links.append('https://www.ruspitomniki.ru' + link)

# print(second_links)

# получаем ссылки на каждую позицию
j = 1
for second_link in second_links[0:4]:
    url = second_link
    response = requests.get(url, headers={'User-Agent': useragent.random})
    print(response.status_code, j)
    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    links_soup = soup.find_all('span', class_='mehrPad')
    for link in links_soup:
        link = link.find('a').get('href')
        third_links.append('https://www.ruspitomniki.ru' + link)

    pictures_soup = soup.find_all('span', class_='imgWr')
    for picture in pictures_soup:
        picture = picture.find('img').get('src')
        pictures.append('https://www.ruspitomniki.ru' + picture)

    j = j + 1

# print(third_links)

name = []
latin_name = []
description = []
table_info = []
photo = []

# идем по каждой ссылке и собираем информацию
j = 1
for third_link in third_links:
    url = third_link
    response = requests.get(url, headers={'User-Agent': useragent.random})
    print(response.status_code, j)
    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    try:
        name_soup = soup.find('div', class_='col-xs-12 col-sm-7').find('h1')
        name_page = name_soup.text
        name.append(name_page)
    except:
        name.append('-')

    try:
        latin_name_soup = soup.find('div', class_='latin')
        latin_name_page = latin_name_soup.text
        latin_name.append(latin_name_page)
    except:
        latin_name.append('-')

    try:
        description_soup = soup.find('div', class_='col-xs-9').find('p')
        description_page = description_soup.text
        description.append(description_page)
    except:
        description.append('-')

    try:
        table_info_soup = soup.find('table', class_='pfBehs')    #.find('tbody')
        table_info_page = table_info_soup.text
        table_info.append(table_info_page)
    except:
        table_info.append('-')


    #     # сплитим текст таблицы, создаем список, из списка создаем словарь
    #     table_list = table_info_page.split('\n')
    #     tab = [x for x in table_list if (x != '') and (x != '\xa0 ')]
    #     tab_dict = {}
    #     for i in range(0, (len(tab) - 1), 2):
    #         tab_dict[tab[i].replace('\\', '').replace('\t', '')] = tab[i + 1].replace('\\', '').replace('\t', '')
    #     table_info.append(tab_dict)
    # except:
    #     table_info.append({})
    j = j + 1

# print(name)
# print(latin_name)
# print(description)
# print(table_info)
# print(pictures)


# Далее записываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('Ruspitomniki')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 2
for i in range(len(name)):
    sheet.write(m, 0, name[i], font)
    sheet.write(m, 1, latin_name[i], font)
    sheet.write(m, 2, third_links[i], font)
    sheet.write(m, 3, pictures[i], font)
    sheet.write(m, 4, description[i], font)
    sheet.write(m, 5, table_info[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 5000  # Ширина колонки
sheet.col(1).width = 5000
sheet.col(2).width = 20000
sheet.col(3).width = 20000
sheet.col(4).width = 25000
sheet.col(5).width = 25000


sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('Ruspitomniki.xls')  # Сохраняем в файл






from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import xlwt
import time
import random
import socks
import socket

start_time = time.time()

useragent = UserAgent()

links = []
years = []
km_age = []
body_type = []
price = []
engine_power = []
transmission = []
owners_count = []
car_name = []

for i in range(1, 101): # 100 страниц
    url = 'https://auto.drom.ru/kia/rio/page' + str(i)
    response = requests.get(url, headers={'User-Agent': useragent.random})  # подменяем агент
    print(response.status_code, i)  # статус сервера (должен быть 200)

    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    links_soup = soup.find('div', class_='css-10ib5jr e93r9u20').find_all('a', class_='css-1hgk7d1 eiweh7o2')
    for link in links_soup:
        try:
            link = link.get('href')
            links.append(link)
        except:
            links.append('none')

# идем по ссылкам всех заказов и забираем требуемую информацию
j = 1
for link in links:
    url = link
    response = requests.get(url, headers={'User-Agent': useragent.random})
    # time.sleep(1)
    print(response.status_code, j)
    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    # Год выпуска
    try:
        # years.append(soup.find('h1', class_='css-cgwg2n e18vbajn0').text)
        year = soup.find('h1', class_='css-cgwg2n e18vbajn0').text
        temp_year = [int(s) for s in year.split() if s.isdigit()]
        years.append(temp_year[0])
    except:
        years.append('None')
    # Мощность
    try:
        engine_power.append(soup.find('div', class_='css-0 epjhnwz1').find('tr', class_='css-10191hq ezjvm5n2')
                            .next_sibling.find('a').text.replace('\xa0л.с.', ''))
    except:
        engine_power.append('None')

    # Тип трансмиссии
    try:
        transmission.append(soup.find('div', class_='css-0 epjhnwz1').find('tr', class_='css-10191hq ezjvm5n2')
                         .next_sibling.next_sibling.next_sibling.find('td').text)
    except:
        transmission.append('None')
    # Тип кузова
    try:
        body_type.append(soup.find('div', class_='css-0 epjhnwz1').find('tr', class_='css-10191hq ezjvm5n2')
                         .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.find('td').text)
    except:
        body_type.append('None')
    # Поколение
    try:
        temp_car_name = soup.find('div', class_='css-0 epjhnwz1').find('tr', class_='css-10191hq ezjvm5n2')\
                         .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling\
                        .next_sibling.next_sibling.next_sibling.next_sibling.find('a').text
        if 'поколение' in temp_car_name:
            car_name.append(temp_car_name)
        else:
            car_name.append('None')
    except:
        car_name.append('None')
    # Пробег
    try:
        temp_km_age = soup.find('div', class_='css-0 epjhnwz1').find('tr', class_='css-10191hq ezjvm5n2')\
                         .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.find('td').text.replace(' ','')
        try:
            km_age.append(int(temp_km_age))
        except:
            km_age.append('Пробег не указан')
    except:
        km_age.append('None')
    # Цена
    try:
        price.append(soup.find('div', class_='css-1hu13v1 e162wx9x0').text.replace('\xa0', '').replace('q', '')
                     .replace(".css-rj9fp4{font-family:'Rouble',sans-serif;}", ''))
    except:
        price.append('None')
    # Количество владельцев
    try:
        owners_count.append(soup.find('table', class_='css-d9xre2 eppj3wm0').find('tr', class_='css-10191hq ezjvm5n2')
                            .next_sibling.find('td').text)
    except:
        owners_count.append('None')
    j = j + 1

# print(links)
# print(years)
# print(engine_power)
# print(transmission)
# print(body_type)
# print(km_age)
# print(car_name)
# print(price)
# print(owners_count)


# Далее зиписываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('KIA Rio info DROM')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 0
for i in range(len(links)):
    sheet.write(m, 0, car_name[i], font)
    sheet.write(m, 1, years[i], font)
    sheet.write(m, 2, km_age[i], font)
    sheet.write(m, 3, body_type[i], font)
    sheet.write(m, 4, engine_power[i], font)
    sheet.write(m, 5, transmission[i], font)
    sheet.write(m, 6, owners_count[i], font)
    sheet.write(m, 7, price[i], font)
    sheet.write(m, 8, links[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 8000  # Ширина колонки
sheet.col(1).width = 3000
sheet.col(2).width = 7000
sheet.col(3).width = 6000
sheet.col(4).width = 3000
sheet.col(5).width = 6000
sheet.col(6).width = 5000
sheet.col(7).width = 6000
sheet.col(8).width = 35000

sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('Kia_Rio_data_set_DROM.xls')  # Сохраняем в файл

print('Количество позиций: ', end='')
print(len(links))
print('Время выполнения: ', end='')
print("%s seconds" % (time.time() - start_time))


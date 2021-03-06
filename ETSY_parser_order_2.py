﻿from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import xlwt
import time
import random
import socks
import socket

target_list = ['https://www.etsy.com/shop/NikkiPattern/reviews?ref=pagination&page=',
               'https://www.etsy.com/shop/CrossStitchingLovers/reviews?ref=pagination&page=',
               'https://www.etsy.com/uk/shop/plasticlittlecovers/reviews?ref=pagination&page=',
               'https://www.etsy.com/shop/VladaXstitch/reviews?ref=pagination&page=',
               'https://www.etsy.com/shop/AlitonEmbroidery/reviews?ref=pagination&page=',
               'https://www.etsy.com/shop/Sewingseed/reviews?ref=pagination&page=',
               'https://www.etsy.com/shop/2x2StitchArt/reviews?ref=pagination&page=',
               'https://www.etsy.com/shop/GentleFeather/reviews?ref=pagination&page=',
               'https://www.etsy.com/ru/shop/PeppermintPurple/reviews?ref=pagination&page=',
               'https://www.etsy.com/shop/Love4CrossStitch/reviews?ref=pagination&page=',
               'https://www.etsy.com/ru/shop/ElCrossStitch/reviews?ref=pagination&page=',
               'https://www.etsy.com/uk/shop/diana70/reviews?ref=pagination&page=',
               'https://www.etsy.com/ru/shop/galabornpatterns/reviews?ref=pagination&page=',
               'https://www.etsy.com/ru/shop/redbeardesign/reviews?ref=pagination&page=',
               'https://www.etsy.com/uk/shop/WellStitches/reviews?ref=pagination&page=',
               'https://www.etsy.com/ru/shop/PineconeMcGee/reviews?ref=pagination&page=']

# работаем под запущенным Tor-браузером
socks.set_default_proxy(socks.SOCKS5, "localhost", 9150)
socket.socket = socks.socksocket

useragent = UserAgent()

for k in range(len(target_list)):
    names = []
    links = []
    images = []
    pages_list = []

    # извлекаем имя магазина и кол-во страниц пагинации
    url = target_list[k] + '1'
    response = requests.get(url, headers={'User-Agent': useragent.random})  # подменяем агент
    if response.status_code != 200:
        print(response.status_code)
        time.sleep(random.randrange(60, 80, 1))
        url = target_list[k] + '1'
        response = requests.get(url, headers={'User-Agent': useragent.random})
        print('  ', response.status_code, k)

    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    shop_name_soup = soup.find('div', class_='flag condensed-header-shop')
    shop_name = shop_name_soup.find('div', class_='hide-xs hide-sm')
    shop_name = shop_name.text.replace('\n', '').strip() # имя магазина

    pagination_soup = soup.find('ul', class_='btn-group-md list-unstyled text-left')
    all_pages = pagination_soup.find_all('li', class_='btn btn-list-item btn-secondary btn-group-item-md hide-xs hide-sm hide-md')

    for page in all_pages:
        page_number = page.find('span', class_='screen-reader-only').text.replace('\n', '').strip().replace('Page ', '')
        pages_list.append(int(page_number))

    pagination_end = max(pages_list) # кол-во страниц пагинации
    print(shop_name)

    # парсим текущий магазин
    for i in range(1, pagination_end + 1):
        url = target_list[k] + str(i)
        response = requests.get(url, headers={'User-Agent': useragent.random})  # подменяем агент
        print(response.status_code, i)  # статус сервера (должен быть 200)
        # если статус не 200, делаем паузу 60-80 сек и снова делаем запрос:
        if response.status_code != 200:
            time.sleep(random.randrange(60, 80, 1))
            url = target_list[k] + str(i)
            response = requests.get(url, headers={'User-Agent': useragent.random})
            print('  ', response.status_code, i)

        html = response.content
        soup = BeautifulSoup(html, 'html.parser')

        names_soup = soup.find_all('div', class_='flag-body hide-xs hide-sm')
        for name in names_soup:
            name = name.find('p').text
            names.append(name)

        links_soup = soup.find_all('div', class_='mt-xs-3 clearfix')
        for link in links_soup:
            try:
                link = link.find('a').get('href')
                links.append('https://www.etsy.com' + link)
            except:
                print('  No link')

        images_soup = soup.find_all('div', class_='card-img-wrap')
        for image in images_soup:
            try:
                image = image.find('img').get('src')
                images.append(image)
            except:
                print('  No image')

    # Далее записываем данные в файл .xls
    book = xlwt.Workbook('utf8')  # Создаем книгу
    # Создаем шрифт
    font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
        italic off; align: wrap on, vert top, horiz left;\
        pattern: pattern solid, fore_colour white;')
    # Добавляем лист
    sheet = book.add_sheet(shop_name)
    # Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
    m = 0
    for i in range(len(names)):
        sheet.write(m, 0, names[i], font)
        sheet.write(m, 1, links[i], font)
        sheet.write(m, 2, images[i], font)
        m = m + 1

    sheet.row(1).height = 2500  # Высота строки
    sheet.col(0).width = 20000  # Ширина колонки
    sheet.col(1).width = 20000
    sheet.col(2).width = 20000

    sheet.portrait = False  # Лист в положении "альбом"
    sheet.set_print_scaling(85)  # Масштабирование при печати
    file_name = 'Shop_' + str(k + 1) + '_' + shop_name + '.xls'
    book.save(file_name)  # Сохраняем в файл

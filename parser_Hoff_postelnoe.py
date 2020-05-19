from selenium import webdriver
from bs4 import BeautifulSoup
import time
import xlwt

links = []
artikul = []
name = []
price = []
price_discount = []

for i in range(1, 3):  # 2 страницы
    driver = webdriver.Chrome(executable_path="C:\\Users\\Alexander\\.PyCharmCE2019.3\\chromedriver.exe")
    driver.get('https://hoff.ru/catalog/tovary_dlya_doma/tekstil/postelnoe_bele/detskoe_postelnoe_bele/page' + str(i))
    time.sleep(2)
    print(i)
    current_page = driver.page_source
    soup = BeautifulSoup(current_page, 'html.parser')
    links_soup = soup.find_all('div', class_='elem-product__name-mobile')
    for link in links_soup:
        link = link.find('a').get('href')
        links.append('https://hoff.ru' + link)
    print(len(links))
    driver.quit()

print(len(links), ' - количество товаров')
# идем по ссылкам всех товаров и забираем требуемую информацию
j = 1
for link in links:
    driver = webdriver.Chrome(executable_path="C:\\Users\\Alexander\\.PyCharmCE2019.3\\chromedriver.exe")
    driver.get(link)
    order_page = driver.page_source
    soup = BeautifulSoup(order_page, 'html.parser')

    try:
        name_soup = soup.find('h1', class_='elem-header__title')
        name_page = list(name_soup.text.strip().split('\n'))
        name.append(name_page[0])
    except:
        name.append('нет заголовка')

    try:
        artikul_soup = soup.find('span', class_='elem-header__articul')
        artikul_item = artikul_soup.text
        artikul.append(artikul_item)
    except:
        artikul.append('-')

    try:
        price_soup = soup.find('div', class_='price-current')
        price_item = price_soup.text.replace('P', '').replace(' ', '')
        price.append(price_item)
    except:
        price.append('товар закончился')

    try:
        price_discount_soup = soup.find('span', class_='price-old')
        price_discount_item = price_discount_soup.text.replace(' ', '')
        price_discount.append(price_discount_item)
    except:
        price_discount.append('-')

    print(j)
    j = j + 1
    time.sleep(1)
    driver.quit()

# Далее зиписываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('Hoff постельное белье')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 0
for i in range(len(links)):
    sheet.write(m, 0, name[i], font)
    sheet.write(m, 1, artikul[i], font)
    # sheet.write(m, 2, brand[i], font)
    sheet.write(m, 3, price[i], font)
    sheet.write(m, 4, price_discount[i], font)
    sheet.write(m, 5, links[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 22000  # Ширина колонки
sheet.col(1).width = 6000
sheet.col(2).width = 500
sheet.col(3).width = 3000
sheet.col(4).width = 3000
sheet.col(5).width = 26000

sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('Hoff_postelnoe.xls')  # Сохраняем в файл

from selenium import webdriver
from bs4 import BeautifulSoup
import time
import xlwt

links = []
artikul = []
name = []
price = []
brand = []

driver = webdriver.Chrome(executable_path="C:\\Users\\Alexander\\.PyCharmCE2019.3\\chromedriver.exe")

for i in range(1, 17):  # 16 страниц
    driver = webdriver.Chrome(executable_path="C:\\Users\\Alexander\\.PyCharmCE2019.3\\chromedriver.exe")
    driver.get('https://market.leroymerlin.ru/catalogue/detskoe-postelnoe-bele/?page=' + str(i))
    time.sleep(1)
    print(i)
    current_page = driver.page_source
    soup = BeautifulSoup(current_page, 'html.parser')
    links_soup = soup.find_all('div', class_='catalog__pic')
    for link in links_soup:
        link = link.find('a').get('href')
        links.append('https://market.leroymerlin.ru' + link)
    print(len(links))
    driver.quit()

print(len(links), ' - количество товаров')
# идем по ссылкам всех товаров и забираем требуемую информацию
j = 1
driver = webdriver.Chrome(executable_path="C:\\Users\\Alexander\\.PyCharmCE2019.3\\chromedriver.exe")
for link in links:
    driver.get(link)
    order_page = driver.page_source
    soup = BeautifulSoup(order_page, 'html.parser')

    try:
        name_soup = soup.find('div', class_='product__header product__header_visible-desktop').find('h1')
        name_page = name_soup.text
        name.append(name_page)
    except:
        name.append('нет заголовка')

    try:
        brand_soup = soup.find('p', class_='product__brand').find('a')
        brand_item = brand_soup.text
        brand.append(brand_item)
    except:
        brand.append('-')

    try:
        price_soup = soup.find('div', class_='product__price-block-row').find('p', class_='product__price')
        price_item = price_soup.text.replace('Цена: \n', '').strip().replace('руб./шт.', '').replace(' ', '')
        price.append(price_item)
    except:
        price.append('товар закончился')

    try:
        artikul_soup = soup.find('p', class_='product__article')
        artikul_item = artikul_soup.text.strip().replace('Артикул: ', '')
        artikul.append(artikul_item)
    except:
        artikul.append('-')

    print(j)
    j = j + 1

driver.quit()

# Далее зиписываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('леруа постельное белье')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 0
for i in range(len(links)):
    sheet.write(m, 0, name[i], font)
    sheet.write(m, 1, artikul[i], font)
    sheet.write(m, 2, brand[i], font)
    sheet.write(m, 3, price[i], font)
    sheet.write(m, 4, links[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 22000  # Ширина колонки
sheet.col(1).width = 3000
sheet.col(2).width = 5000
sheet.col(3).width = 3000
sheet.col(4).width = 26000

sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('leroy_merlin_postelnoe.xls')  # Сохраняем в файл

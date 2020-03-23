from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import re
import xlwt
import random

links = []
name = []
email = []

# запускаем chrome с нашим профилем пользователя
file_name_profile = r'C:\\Users\\user1174\\AppData\\Local\\Google\\Chrome\\User Data'
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=" + file_name_profile)
driver = webdriver.Chrome(executable_path="C:\\Users\\user1174\\.PyCharmCE2019.3\\chromedriver.exe", options=options)
driver.get("https://www.inkydeals.com/wp-admin/admin.php?page=wcpv-vendor-orders")

# парсим ссылки на заказы с первой страницы
first_page = driver.page_source
soup = BeautifulSoup(first_page, 'html.parser')
links_soup = soup.find_all('a', attrs={'class': 'wcpv-vendor-order-by-id'})
links = [link.attrs['href'] for link in links_soup]
time.sleep(random.randrange(3, 8, 1))

# парсим ссылки на заказы с последующих страниц (2 - 65)
for i in range(2, 66):
    driver.get("https://www.inkydeals.com/wp-admin/admin.php?page=wcpv-vendor-orders&paged=" + str(i))
    current_page = driver.page_source
    soup = BeautifulSoup(current_page, 'html.parser')
    links_soup = soup.find_all('a', attrs={'class': 'wcpv-vendor-order-by-id'})
    links += [link.attrs['href'] for link in links_soup]
    time.sleep(random.randrange(3, 8, 1))

time.sleep(random.randrange(3, 8, 1)) # задержка времени от 3 до 7 сек
# идем по ссылкам всех заказов и забираем требуемую информацию
for link in links:
    driver.get(link)
    order_page = driver.page_source
    soup = BeautifulSoup(order_page, 'html.parser')
    email += soup.find('div', class_='address').find('a')  # берем адреса эл.почты
    name_string = str((soup.find('div', class_='address').find('p'))) # получаем грязную строку с именем
    clean_name = re.findall(r'</strong>(.*?)<', name_string, re.DOTALL) # выдергиваем имя регуляркой
    name = name + clean_name
    time.sleep(random.randrange(3, 8, 1))

driver.quit()

# Далее зиписываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('список адресов')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 0
for i in range(len(email)):
    sheet.write(m, 0, name[i], font)
    sheet.write(m, 1, email[i], font)
    sheet.write(m, 2, links[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 12000  # Ширина колонки
sheet.col(1).width = 12000
sheet.col(2).width = 25000
sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('data_Pavel.xls')  # Сохраняем в файл

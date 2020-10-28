from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import xlwt
import random

start_time = time.time()

links = []
years = []
km_age = []
body_type = []
price = []
engine_power = []
transmission = []
condition = []
owners_count = []
car_name = []

# запускаем chrome с нашим профилем пользователя
file_name_profile = r'C:\\Users\\user1174\\AppData\\Local\\Google\\Chrome\\User Data'
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=" + file_name_profile)
driver = webdriver.Chrome(executable_path="C:\\Users\\user1174\\.PyCharmCE2019.3\\chromedriver.exe", options=options)

for i in range(1, 101): # 100 страниц
    url = 'https://www.avito.ru/rossiya/avtomobili/kia/rio-ASgBAgICAkTgtg3KmCjitg3Krig?cd=1&p=' + str(i)
    driver.get(url)
    current_page = driver.page_source
    soup = BeautifulSoup(current_page, 'html.parser')

    links_soup = soup.find('div', class_='index-root-2c0gs').find_all('a', class_='item-slider item-slider--4-3')
    for link in links_soup:
        try:
            link = link.get('href')
            links.append('https://www.avito.ru' + link)
        except:
            links.append('none')
    if i % 10 == 0:
        time.sleep(random.randrange(10, 20, 1))
    print(i)

# идем по ссылкам всех машин и забираем требуемую информацию
j = 1
for link in links:
    driver.get(link)
    current_page = driver.page_source
    soup = BeautifulSoup(current_page, 'html.parser')
    # Поколение
    try:
        car_name.append(soup.find('ul', class_='item-params-list').find('li', class_='item-params-list-item')
                         .next_sibling.next_sibling.next_sibling.next_sibling.text.replace(' Поколение: ', '')
                        .split('(')[0].strip())
    except:
        car_name.append('None')
    # Мощность
    try:
        power = soup.find('ul', class_='item-params-list').find('li', class_='item-params-list-item')\
                          .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.text
        temp_power = [int(s) for s in power.split('(')[1].split() if s.isdigit()]
        engine_power.append(temp_power[0])

    except:
        engine_power.append('None')
    # Год выпуска
    try:
        years.append(soup.find('ul', class_='item-params-list').find('li', class_='item-params-list-item')\
                          .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                        .next_sibling.next_sibling.next_sibling.text.replace(' Год выпуска: ', '').strip())
    except:
        years.append('None')
    # Пробег
    try:
        km_age_temp = soup.find('ul', class_='item-params-list').find('li', class_='item-params-list-item')\
                          .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling\
                         .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.text\
                      .replace(' Пробег: ', '').replace('\xa0км ', '')
        if len(km_age_temp) < 8:
            km_age.append(km_age_temp)
        else:
            km_age.append('None')
    except:
        km_age.append('None')
    # Состояние
    try:
        condition.append(soup.find('ul', class_='item-params-list').find('li', class_='item-params-list-item') \
                                   .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                          .next_sibling.next_sibling.next_sibling.next_sibling
                         .next_sibling.next_sibling.next_sibling.text.replace(' Состояние: ', '').strip())
    except:
        condition.append('None')
    # Количество владельцев
    try:
        owners_count.append(soup.find('ul', class_='item-params-list').find('li', class_='item-params-list-item') \
                                   .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                          .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                         .next_sibling.next_sibling.next_sibling.text.replace(' Владельцев по ПТС: ', '').strip())
    except:
        owners_count.append('None')
    # Тип кузова
    try:
        body_type.append(soup.find('ul', class_='item-params-list').find('li', class_='item-params-list-item') \
                                   .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                          .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                         .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                         .text.replace('Тип кузова: ', '').strip())
    except:
        body_type.append('None')
    # Тип трансмиссии
    try:
        transmission.append(soup.find('ul', class_='item-params-list').find('li', class_='item-params-list-item') \
                                   .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                          .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                            .next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling
                            .next_sibling.next_sibling.next_sibling.next_sibling
                         .next_sibling.next_sibling.next_sibling.text.replace(' Коробка передач: ', '').strip())
    except:
        transmission.append('None')
    # Цена
    try:
        price.append(soup.find('div', class_='item-price').find('span', class_='js-item-price').text.replace(' ', ''))
    except:
        price.append('None')

    if j % 20 == 0:
        time.sleep(random.randrange(10, 20, 1))
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
sheet = book.add_sheet('KIA Rio info AVITO')
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
    sheet.write(m, 7, condition[i], font)
    sheet.write(m, 8, price[i], font)
    sheet.write(m, 9, links[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 8000  # Ширина колонки
sheet.col(1).width = 3000
sheet.col(2).width = 7000
sheet.col(3).width = 6000
sheet.col(4).width = 3000
sheet.col(5).width = 6000
sheet.col(6).width = 5000
sheet.col(7).width = 8000
sheet.col(8).width = 6000
sheet.col(9).width = 35000

sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('Kia_Rio_data_set_Avito.xls')  # Сохраняем в файл

print('Количество позиций: ', end='')
print(len(links))
print('Время выполнения: ', end='')
print("%s seconds" % (time.time() - start_time))

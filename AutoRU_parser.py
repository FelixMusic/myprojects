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
color = []
price = []
engine_volume = []
engine_power = []
fuel_type = []
transmission = []
drive = []
wheel = []
condition = []
owners_count = []
passport = []
customs = []
complectation_type = []
region = []
car_name = []


# запускаем chrome с нашим профилем пользователя
file_name_profile = r'C:\\Users\\user1174\\AppData\\Local\\Google\\Chrome\\User Data'
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=" + file_name_profile)
driver = webdriver.Chrome(executable_path="C:\\Users\\user1174\\.PyCharmCE2019.3\\chromedriver.exe", options=options)


for i in range(1, 3): # 35 страниц Kia Rio
    driver.get('https://auto.ru/moskva/cars/kia/rio/all/?page=' + str(i) + '&output_type=list&geo_id=21656')
    current_page = driver.page_source
    soup = BeautifulSoup(current_page, 'html.parser')
    links_soup = soup.find_all('a', attrs={'class': 'Link ListingItemTitle-module__link'})
    links += [link.attrs['href'] for link in links_soup]
    time.sleep(random.randrange(2, 5, 1))


# идем по ссылкам всех автомобилей и забираем требуемую информацию
for link in links:
    driver.get(link)
    car_page = driver.page_source
    soup = BeautifulSoup(car_page, 'html.parser')
    # Год выпуска
    try:
        years.append(soup.find('li', class_='CardInfo__row CardInfo__row_year').find('a').text)
    except:
        years.append('None')

    # Пробег
    try:
        km_age.append(soup.find('li', class_='CardInfo__row CardInfo__row_kmAge')
                      .find('span').next_sibling.text.replace('\xa0', '').replace('км', ''))
    except:
        km_age.append('None')

    # Тип кузова
    try:
        body_type.append(soup.find('li', class_='CardInfo__row CardInfo__row_bodytype').find('span').next_sibling.text)
    except:
        body_type.append('None')

    # Цвет
    try:
        color.append(soup.find('li', class_='CardInfo__row CardInfo__row_color').find('span').next_sibling.text)
    except:
        color.append('None')

    # Цена
    try:
        price.append(soup.find('span', class_='OfferPriceCaption__price').text.replace('\xa0', '').replace('₽', ''))
    except:
        price.append('None')

    # Объем двигателя, мощность и тип топлива
    try:
        engine_row = soup.find('li', class_='CardInfo__row CardInfo__row_engine').find('div').text
        engine_list = engine_row.split('/')

        engine_volume.append(engine_list[0].replace(' л', ''))
        engine_power.append(engine_list[1].replace('\xa0л.с.', '').strip())
        fuel_type.append(engine_list[2])
    except:
        engine_volume.append('None')
        engine_power.append('None')
        fuel_type.append('None')

    # Тип трансмисии
    try:
        transmission.append(soup.find('li', class_='CardInfo__row CardInfo__row_transmission').find('span').next_sibling.text)
    except:
        transmission.append('None')

    # Тип привода
    try:
        drive.append(soup.find('li', class_='CardInfo__row CardInfo__row_drive').find('span').next_sibling.text)
    except:
        drive.append('None')

    # Расположение руля
    try:
        wheel.append(soup.find('li', class_='CardInfo__row CardInfo__row_wheel').find('span').next_sibling.text)
    except:
        wheel.append('None')

    # Состояние
    try:
        condition.append(soup.find('li', class_='CardInfo__row CardInfo__row_state').find('span').next_sibling.text)
    except:
        condition.append('None')

    # Количество владельцев
    try:
        owners_count.append(soup.find('li', class_='CardInfo__row CardInfo__row_ownersCount').find('span').next_sibling.text)
    except:
        owners_count.append('None')

    # Информация о ПТС
    try:
        passport.append(soup.find('li', class_='CardInfo__row CardInfo__row_pts').find('span').next_sibling.text)
    except:
        passport.append('None')

    # Сведения о таможне
    try:
        customs.append(soup.find('li', class_='CardInfo__row CardInfo__row_customs').find('span').next_sibling.text)
    except:
        customs.append('None')

    # Тип комплектации
    try:
        complectation_type.append(soup.find('div', class_='CardComplectation__titleWrap').find('h2').text)
    except:
        complectation_type.append('None')

    # Регион продажи
    try:
        region.append(soup.find('span', class_='MetroListPlace__regionName MetroListPlace_nbsp').text)
    except:
        region.append('None')

    # Название автомобиля
    try:
        car_name.append(soup.find('div', class_='CardHead-module__title').text)
    except:
        car_name.append('None')

driver.quit()

# Далее зиписываем данные в файл .xls
book = xlwt.Workbook('utf8')  # Создаем книгу
# Создаем шрифт
font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
    italic off; align: wrap on, vert top, horiz left;\
    pattern: pattern solid, fore_colour white;')
# Добавляем лист
sheet = book.add_sheet('KIA Rio info')
# Заполняем ячейки (Строка, Колонка, Текст, Шрифт)
m = 0
for i in range(len(links)):
    sheet.write(m, 0, car_name[i], font)
    sheet.write(m, 1, years[i], font)
    sheet.write(m, 2, km_age[i], font)
    sheet.write(m, 3, body_type[i], font)
    sheet.write(m, 4, color[i], font)
    sheet.write(m, 5, engine_volume[i], font)
    sheet.write(m, 6, engine_power[i], font)
    sheet.write(m, 7, fuel_type[i], font)
    sheet.write(m, 8, transmission[i], font)
    sheet.write(m, 9, drive[i], font)
    sheet.write(m, 10, wheel[i], font)
    sheet.write(m, 11, condition[i], font)
    sheet.write(m, 12, passport[i], font)
    sheet.write(m, 13, customs[i], font)
    sheet.write(m, 14, complectation_type[i], font)
    sheet.write(m, 15, region[i], font)
    sheet.write(m, 16, price[i], font)
    sheet.write(m, 17, links[i], font)
    m = m + 1

sheet.row(1).height = 2500  # Высота строки
sheet.col(0).width = 8000  # Ширина колонки
sheet.col(1).width = 3000
sheet.col(2).width = 3000
sheet.col(3).width = 6000
sheet.col(4).width = 6000
sheet.col(5).width = 3000
sheet.col(6).width = 3000
sheet.col(7).width = 4000
sheet.col(8).width = 6000
sheet.col(9).width = 6000
sheet.col(10).width = 4000
sheet.col(11).width = 8000
sheet.col(12).width = 4000
sheet.col(13).width = 5000
sheet.col(14).width = 12000
sheet.col(15).width = 6000
sheet.col(16).width = 6000
sheet.col(17).width = 35000

sheet.portrait = False  # Лист в положении "альбом"
sheet.set_print_scaling(85)  # Масштабирование при печати
book.save('Kia_Rio_data_set.xls')  # Сохраняем в файл

print('Количество позиций: ', end='')
print(len(links))
print('Время выполнения: ', end='')
print("%s seconds" % (time.time() - start_time))

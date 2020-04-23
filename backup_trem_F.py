import shutil
import datetime
import os
import time

start_time = time.time()


def last_fld():
    date_fld = [i for i in os.walk(path)][0][1]
    last = '00.00.0000'
    for date in date_fld:
        if int(date[6:]) > int(last[6:]):
            last = date
        elif int(date[6:]) == int(last[6:]):
            if int(date[3:5]) > int(last[3:5]):
                last = date
            elif int(date[3:5]) == int(last[3:5]):
                if int(date[0:2]) >= int(last[0:2]):
                    last = date
    return last


path = 'F:\\Дорогов'
deleting_folder = last_fld()
print('Последнее сохранение производилось: ', deleting_folder)

try:
    deleting_dir_path = 'F:\\Дорогов\\' + deleting_folder + '\\VERICUT рабочая'
    shutil.rmtree(deleting_dir_path)
except:
    print('Удаление папки F:\Дорогов\\', deleting_folder, '\VERICUT рабочая - не требуется', sep='')

today = datetime.datetime.today()
datefld = today.strftime('%d.%m.%Y')

print('Резервное копирование началось')
print('Дождитесь сообщения о завершении копирования:')

shutil.copytree('C:\\Users\\user1174\\Desktop\\VERICUT рабочая',
                'F:\\Дорогов\\' + datefld + '\\VERICUT рабочая')
shutil.copytree('C:\\Users\\user1174\\Desktop\\Чертежи заготовок для импорта',
                'F:\\Дорогов\\' + datefld + '\\Чертежи заготовок для импорта')
shutil.copytree('C:\\Users\\user1174\\Desktop\\НУЖНЫЕ ПРИМЕРЫ ПРОГРАММ',
                'F:\\Дорогов\\' + datefld + '\\НУЖНЫЕ ПРИМЕРЫ ПРОГРАММ')
shutil.copytree('\\\KUPAVNA-DATA\\Kupavna\\==ПРОГРАММЫ КУПАВНА==\\ТОКАРНЫЕ ПРОГРАММЫ\\ПРОГРАММИСТЫ\\COLLECTOR_T',
                'F:\\Дорогов\\' + datefld + '\\COLLECTOR_T')
shutil.copytree('\\\KUPAVNA-DATA\\Kupavna\\=ПРОГРАММЫ ЛЮБЕРЦЫ=\\COLLECTOR_F',
                'F:\\Дорогов\\' + datefld + '\\COLLECTOR_F')
shutil.copytree('\\\KUPAVNA-DATA\\Kupavna\\==ПРОГРАММЫ TREM MILL==',
                'F:\\Дорогов\\' + datefld + '\\==ПРОГРАММЫ TREM MILL==')

print('Резервное копирование завершено!')
print('Время выполнения: ', end='')
print("%s seconds" % (time.time() - start_time))
print('Введите любую цифру')
f = int(input())

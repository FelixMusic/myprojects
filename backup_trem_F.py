import shutil
import datetime

now = datetime.datetime.now()

d = list(map(str, str(now).split()))
e = list(map(str, str(d[0]).split('-')))
datefld = str(e[2]) + '.' + str(e[1]) + '.' + str(e[0])

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
print('Введите любую цифру')
f = int(input())

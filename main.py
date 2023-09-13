import csv
import subprocess
import os
import time
import pandas as pd
import openpyxl

list_pc = []
def csv_read():
    with open('laptop.csv', 'r', newline='') as file:
        reader = csv.reader(file)
        for row in reader:
            list_pc.append(row[0])



name_laptop = []
model_laptop = []
Authorized_user = []
All_users = []
emails = []

laptop_info = {'Name Laptop':name_laptop, 'Model Laptop':model_laptop, 'Autorized user':Authorized_user, 'All users':All_users, 'E-mail':emails}

def laptop_inform():
    csv_read()
    for pc in list_pc:
        print(' Опрашиваю ' + pc)
        model = subprocess.getoutput(f'powershell (Get-WmiObject -class Win32_ComputerSystem -ComputerName {pc}).model')
        if model[:3] == 'Get':
            model = 'Нет данных'
        user = subprocess.getoutput(f'powershell (Get-WmiObject -class Win32_ComputerSystem –ComputerName {pc}).Username')
        if user[:3] == 'Get':
            user = 'Нет данных'
        user = user[user.find('\\') + 1:]

        email = subprocess.getoutput(f'powershell Get-ADUser -Identity belidmim -Properties mail')
        a = email.find('@')
        b = email.find(':')
        a = email[a - 40:a + 12]
        b = a.find(':')
        email = a[b + 2:].replace(' ', '')

        users_pc = []
        try:
            content = os.listdir(f'//{pc}/c$/users')
            for i in content:
                if i[:5].lower() == 'admin':
                    continue
                else:
                    users_pc.append(i)
            spis = ['All Users', 'Default', 'Default User', 'desktop.ini', 'Public', 'Администратор', 'Все пользователи']
            for i in spis:
                users_pc.remove(i)
        except:
            users_pc.append('Нет данных')

        name_laptop.append(pc)
        model_laptop.append(model)
        Authorized_user.append(user)
        All_users.append(users_pc)
        emails.append(email
                      )
def check_admin():
    check_user = subprocess.getoutput('whoami')
    check_admin = check_user[check_user.find('\\')+1:]
    if check_admin[:5] != 'admin':
        print('Программа запущена не от имени доменного администратора!')
    else:
        return True
def check_file():
    dir = 'laptop.csv'.lower()
    if os.path.exists(dir):
        return True
    else:
        print('Отсутствует файл ', dir)

print('''
    Для работы данного скрипта необходимо: 
    1) Создать файл laptop.csv рядом с программой
       Количество опрашиваемых PC неограниено
       Файл laptop.csv должен выглядеть так:
    ----------------------------------------------   
    OFSCR122T3
    OFS1BJ3XT2
    OFSH03R0T3
    OFSBDL02T3
    ----------------------------------------------
    Без пробелов, запятых и прочего.
    
    Вывод отчета будет в файл laptop_info.xlsx
    ''')

laptop_inform()


'''
if check_file() and check_admin():
    laptop_inform()
    df = pd.DataFrame(laptop_info)
    df.to_excel('laptop_info.xlsx')
else:
    time.sleep(100)'''

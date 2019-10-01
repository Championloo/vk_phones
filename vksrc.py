import vk
import sqlite3
import xlwt
from tqdm import tqdm
import time
import os

token = ''
try:
	with open ('token', 'r') as f:
		token = f.readline()
except:
	print("""\nДля создания токена перейдите на https://vk.com/editapp?act=create
Введите название, выберите платформу "Standalone-приложение", нажмите "Подключить приложение"
Перейдите в настройки и скопируйте "Сервисный ключ доступа" - это и есть токен.""")
	token = input('\nВведите токен: ')
	with open ('token', 'w') as f:
		f.write(token)
		print('Токен записан')
print('Чтобы изменить токен, откройте файл token с помощью блокнота и перепишите значение или удалите файл\n')
groupid = input('Введите ID или короткое имя (из ссылки) группы/паблика: ').replace('https://vk.com/','')
session = vk.Session(access_token=token)
vk_api = vk.API(session)
if groupid.isdigit()==False:
	groupid = vk_api.groups.getById(group_id=groupid, v=5.92)[0]['id']
con = sqlite3.connect("{}.db".format(groupid))
cursor = con.cursor()

with con:
	cursor.execute("""CREATE TABLE if not exists vk (id int, name text, last text, phone text, city text)""")

first = vk_api.groups.getMembers(group_id=groupid, v=5.92) 
data = first["items"]  
count = first["count"] // 1000  
for i in range(1, count+1):  
	data = data + vk_api.groups.getMembers(group_id=groupid, v=5.92, offset=i*1000)["items"]
with con:
	for d in data:
		cursor.execute("INSERT INTO vk VALUES (?,?,?,?,?)",(str(d),'','','',''))
	con.commit()

with con:
	cursor.execute("SELECT distinct * FROM vk")  
	data = cursor.fetchall()
	for i in tqdm(range(len(data)), smoothing=0, miniters=None):
		try:
			user = vk_api.users.get(user_ids=str(data[i][0]), v=5.92, fields='contacts, city', lang='ru')[0]
			phone = ''
			city = ''
			if 'is_closed' in user and user['is_closed']==False:
				if 'home_phone' in user and len(user['home_phone'])>8:
					hphone = user['home_phone']
					hphone = hphone.lower().replace('один', '1').replace('два', '2').replace('три', '3').replace('четыре', '4').replace('o','0')
					hphone = hphone.lower().replace('пять', '5').replace('шесть', '6').replace('семь', '7').replace('о','0').replace('l','1')
					hphone = hphone.lower().replace('восемь', '8').replace('девять', '9').replace('ноль', '0').replace('з','3').replace('i','1')
					truephone = ''
					for d in hphone:
						if d.isdigit()==True:
							truephone += d
					if len(truephone)>9:
						phone += truephone+'; '
				if 'mobile_phone' in user and len(user['mobile_phone'])>8:
					mphone = user['mobile_phone']
					mphone = mphone.lower().replace('один', '1').replace('два', '2').replace('три', '3').replace('четыре', '4').replace('o','0')
					mphone = mphone.lower().replace('пять', '5').replace('шесть', '6').replace('семь', '7').replace('о','0').replace('l','1')
					mphone = mphone.lower().replace('восемь', '8').replace('девять', '9').replace('ноль', '0').replace('з','3').replace('i','1')
					truephone = ''
					for d in mphone:
						if d.isdigit()==True:
							truephone += d
					if len(truephone)>9:
						phone += truephone+'; '
				if 'city' in user:
					city = user['city']['title']
			cursor.execute("UPDATE vk SET name=?, last=?, phone=?, city=? where id=?",(user['first_name'], user['last_name'], phone, city, str(data[i][0])))
			con.commit()
			time.sleep(0.001)
		except:
			time.sleep(15)
			user = vk_api.users.get(user_ids=str(data[i][0]), v=5.92, fields='contacts, city', lang='ru')[0]
			phone = ''
			city = ''
			if 'is_closed' in user and user['is_closed']==False:
				if 'home_phone' in user and len(user['home_phone'])>8:
					hphone = user['home_phone']
					hphone = hphone.lower().replace('один', '1').replace('два', '2').replace('три', '3').replace('четыре', '4').replace('o','0')
					hphone = hphone.lower().replace('пять', '5').replace('шесть', '6').replace('семь', '7').replace('о','0').replace('l','1')
					hphone = hphone.lower().replace('восемь', '8').replace('девять', '9').replace('ноль', '0').replace('з','3').replace('i','1')
					truephone = ''
					for d in hphone:
						if d.isdigit()==True:
							truephone += d
					if len(truephone)>9:
						phone += truephone+'; '
				if 'mobile_phone' in user and len(user['mobile_phone'])>8:
					mphone = user['mobile_phone']
					mphone = mphone.lower().replace('один', '1').replace('два', '2').replace('три', '3').replace('четыре', '4').replace('o','0')
					mphone = mphone.lower().replace('пять', '5').replace('шесть', '6').replace('семь', '7').replace('о','0').replace('l','1')
					mphone = mphone.lower().replace('восемь', '8').replace('девять', '9').replace('ноль', '0').replace('з','3').replace('i','1')
					truephone = ''
					for d in mphone:
						if d.isdigit()==True:
							truephone += d
					if len(truephone)>9:
						phone += truephone+'; '
				if 'city' in user:
					city = user['city']['title']
			cursor.execute("UPDATE vk SET name=?, last=?, phone=?, city=? where id=?",(user['first_name'], user['last_name'], phone, city, str(data[i][0])))
			con.commit()
			time.sleep(0.001)

with con:
	cursor.execute("SELECT distinct * FROM vk where phone!=''")  
	data = cursor.fetchall()
wb = xlwt.Workbook()
style = xlwt.XFStyle()
wb.add_sheet('Лист1', cell_overwrite_ok=True)
ws = wb.get_sheet('Лист1')
ws.col(0).width = 5000
ws.col(1).width = 10000
ws.col(2).width = 7000
ws.col(3).width = 7000
ws.write(0, 0, 'Имя', style)
ws.write(0, 1, 'Телефон', style)
ws.write(0, 2, 'Ссылка на страницу (id)', style)
ws.write(0, 3, 'Город', style)

for d in range(1, len(data)):
	if ''!=data[d][3]:
		print(data[d])
		user = data[d]
		ws.write(d, 0, user[1]+' '+user[2], style)
		ws.write(d, 1, user[3], style)
		ws.write(d, 2, 'https://vk.com/id'+str(user[0]), style)
		ws.write(d, 3, user[4], style)
wb.save('{}.xls'.format(groupid))
con.close()
os.remove("{}.db".format(groupid))
print()
input('Таблица создана, для выхода нажмите ENTER')

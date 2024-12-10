#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
from mesto import ul
import json
import random
import time
import os
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'zemm/0001-0080_00_У_001-0009_VSE42.xlsx')

class Brsn_Zem(Spider):
     def prepare(self):
	  #self.f = page
	  self.ws = workbook.add_worksheet()
	  self.ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	  self.ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	  self.ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
	  self.ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	  self.ws.write(0, 4, u"УЛИЦА")
	  self.ws.write(0, 5, u"ДОМ")
	  self.ws.write(0, 6, u"ОРИЕНТИР")
	  self.ws.write(0, 7, u"ТРАССА")
	  self.ws.write(0, 8, u"УДАЛЕННОСТЬ")
	  self.ws.write(0, 9, u"ОПЕРАЦИЯ")
	  self.ws.write(0, 10, u"СТОИМОСТЬ")
	  self.ws.write(0, 11, u"ЦЕНА_ЗА_СОТКУ")
	  self.ws.write(0, 12, u"ПЛОЩАДЬ")
	  self.ws.write(0, 13, u"КАТЕГОРИЯ_ЗЕМЛИ")
	  self.ws.write(0, 14, u"ВИД_РАЗРЕШЕННОГО_ИСПОЛЬЗОВАНИЯ")
	  self.ws.write(0, 15, u"ГАЗОСНАБЖЕНИЕ")
	  self.ws.write(0, 16, u"ВОДОСНАБЖЕНИЕ")
	  self.ws.write(0, 17, u"КАНАЛИЗАЦИЯ")
	  self.ws.write(0, 18, u"ЭЛЕКТРОСНАБЖЕНИЕ")
	  self.ws.write(0, 19, u"ТЕПЛОСНАБЖЕНИЕ")
	  self.ws.write(0, 20, u"ОХРАНА")
	  self.ws.write(0, 21, u"ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	  self.ws.write(0, 22, u"ОПИСАНИЕ")
	  self.ws.write(0, 23, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 24, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 25, u"ТЕЛЕФОН")
	  self.ws.write(0, 26, u"КОНТАКТНОЕ_ЛИЦО")
	  self.ws.write(0, 27, u"КОМПАНИЯ")
	  self.ws.write(0, 28, u"ДАТА_РАЗМЕЩЕНИЯ")
	  self.ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
	  self.ws.write(0, 30, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(1,11):
               yield Task ('post',url='http://dom.vse42.ru/property/lands/page%d'%x,network_try_count=100)
                 
     def task_post(self,grab,task):
    
	  for elem in grab.doc.select(u'//a[@class="obj-adv-url"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	  
        
        
     def task_item(self, grab, task):
	  try:
	       sub = u'Кемеровская Область'#grab.doc.select(u'//label[contains(text(),"Адрес:")]/following-sibling::p').text().split(', ')[0]
	  except IndexError:
	       sub = ''
	  try:
	       ray = grab.doc.select(u'//span[@class="beforeheader"]').text()
	     #print ray 
	  except DataNotFound:
	       ray = ''          
	  try:
	       punkt= u'Кемерово'#grab.doc.select(u'//span[@class="hp_caption"][contains(text(),"Город:")]/following-sibling::text()').text()
	  except IndexError:
	       punkt = ''
	       
	  try:
	       ter= grab.doc.select(u'//span[@class="beforeheader"]').text().split(u'Кемерово ')[1]
	  except IndexError:
	       ter =''
	       
	  try:
	       uli = grab.doc.select(u'//h1[@class="title"]').text()
	       for x in range(len(ul)):
	            if ul[x] in uli:
		         uls = ul[x]+uli.split(ul[x])[1]
	       uliza = uls.split(', ')[0]
	  except IndexError:
	       uliza = ''
	       
          try:
               dmk = grab.doc.select(u'//h1[@class="title"]').text()
               for x in range(len(ul)):
	            if ul[x] in dmk:
	                 dm = dmk.split(ul[x])[1]
               dom = dm.split(', ')[1]
          except IndexError:
               dom = ''
	       
	  try:
	       trassa = grab.doc.select(u'//label[contains(text(),"Шоссе:")]/following-sibling::p').text().split(', ')[0]
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//label[contains(text(),"Растояние до города:")]/following-sibling::span').text()
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//div[contains(text(),"Цена")]/following-sibling::div[1]').text()
	  except IndexError:
	       price = ''
	  try:
	       plosh = grab.doc.select(u'//div[contains(text(),"Площадь участка")]/following-sibling::div[1]').text()
	  except IndexError:
	       plosh = ''
	  try:
	       vid = grab.doc.select(u'//div[contains(text(),"Тип дома")]/following-sibling::div[1]').text()
	  except DataNotFound:
	       vid = '' 
	       
	       
	  #try:
	       #ohrana = grab.doc.select(u'//h3').text().split(u'Цена - ')[0].split(u'Участки, ')[1].replace(vid,'')[2:][0:-2]
	  #except DataNotFound:
	  ohrana =''
	  try:
	       gaz = grab.doc.select(u'//label[contains(text(),"Газ:")]/following-sibling::span').text()
	  except DataNotFound:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//label[contains(text(),"Вода:")]/following-sibling::span').text()
	  except DataNotFound:
	       voda =''
	  try:
	       kanal = re.sub(u'^.*(?=санузел)','', grab.doc.select(u'//*[contains(text(), "санузел")]').text())[:7].replace(u'санузел',u'есть')
	  except DataNotFound:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//label[contains(text(),"Свет:")]/following-sibling::span').text()
	  except DataNotFound:
	       elek =''
	       
	  teplo =''
	       
	  try:
	       oper = u'Продажа'#grab.doc.select(u'//label[contains(text(),"Тип операции:")]/following-sibling::span').text().replace(u'Сдам',u'Аренда').replace(u'Спрос',u'Аренда').replace(u'Продам',u'Продажа')  
	  except DataNotFound:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//h3[contains(text(),"Дополнительная информация")]/following-sibling::div[1]').text() 
	  except DataNotFound:
	       opis = ''
	       
	 
	 
	       
	  try:
	       comp = grab.doc.select(u'//div[contains(text(),"Опубликовал")]/following-sibling::div[1]').text()
	  except IndexError:
	       comp = ''
	       
	  try:
	       conv = [(u' Августа ',u'.08.'), (u' Июля ',u'.07.'),
	            (u' Мая ',u'.05.'),(u' Июня ',u'.06.'),
	            (u' Марта ',u'.03.'),(u' Апреля ',u'.04.'),
	            (u' Января ',u'.01.'),(u' Декабря ',u'.12.'),
	            (u' Сентября ',u'.09.'),(u' Ноября ',u'.11.'),
	            (u' Февраля ',u'.02.'),(u' Октября ',u'.10.')] 	       
	       dt= grab.doc.select(u'//span[contains(text(),"Создано:")]').text().split(u'Создано: ')[1]
	       data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
	  except IndexError:
	       data = ''
	       
	  try:       
	       dt1= grab.doc.select(u'//span[contains(text(),"Обновленно:")]').text().split(u'Обновленно: ')[1]
	       if 'дней' in dt1:
	            lico = '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=int(re.sub(u'[^\d]','',dt1))))
	       elif 'день' in dt1:
	            lico = '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=int(re.sub(u'[^\d]','',dt1))))
	       elif 'дня' in dt1:
	            lico = '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=int(re.sub(u'[^\d]','',dt1))))
	       else:
	            lico = datetime.today().strftime('%d.%m.%Y')	       
	  except IndexError:
	       lico = ''	       
	       
	       
	       
		    
	  
						   
	       
	  projects = {'url': task.url,
                      'sub': sub,
                      'rayon': ray,
                      'punkt': punkt,
                      'teritor': ter,
                      'ulica': uliza.replace(u'улица не указана',''),
	              'dom': dom,
                      'trassa': trassa,
                      'udal': udal,
                      'cena': price,
                      'plosh':plosh,
                      'vid': vid,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
	              'phone':random.choice(list(open('../phone.txt').read().splitlines())),
                      'lico':lico,
                      'company':comp,
                      'data':data,
                      'oper':oper
                      }
	  
	  #try:
	       #link =  grab.make_url_absolute(grab.doc.select(u'//a[@class="get_contacts"]').attr('data-href'))
	       #headers ={'Accept': '*/*',
			 #'Accept-Encoding': 'gzip,deflate',
			 #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			 #'Cookie': 'PHPSESSID = 3hsdpd3627c0jmdidhik9afo00',
			 #'Host': 'dom.vse42.ru',
			 #'Referer': task.url,
			 #'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0', 
			 #'X-Requested-With' : 'XMLHttpRequest'}
	       ##print link
	       #gr = Grab()
	       #gr.setup(url=link,headers=headers)
	       #yield Task('phone',grab=gr,project=projects,refresh_cache=True,network_try_count=100)
	  #except IndexError:
	       #yield Task('phone',grab=grab,project=projects)	  
	  
     #def task_phone(self, grab, task):
	  ##json_data = json.loads(grab.response.body)
	  #try:
	       #phone = re.sub('[^\d\+]','',grab.response.json['text'])
	  #except (IndexError,ValueError):
	       #phone= random.choice(list(open('../phone.txt').read().splitlines()))
	       
	       
	       
          
	  #yield Task('write',project=task.project,phone=phone,grab=grab)
	  yield Task('write',project=projects,grab=grab)
            
     def task_write(self,grab,task):
	  print('*'*50)
	  print  task.project['sub']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['teritor']
	  print  task.project['ulica']
	  print  task.project['dom']
	  print  task.project['trassa']
	  print  task.project['udal']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['vid']
	  print  task.project['ohrana']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanaliz']
	  print  task.project['electr']
	  print  task.project['teplo']
	  print  task.project['opis']
	  print task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	  print  task.project['oper']
	  
	  #global result
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 31, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 3, task.project['teritor'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 7, task.project['trassa'])
	  self.ws.write(self.result, 8, task.project['udal'])
	  self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 14, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  #self.ws.write(self.result, 29, task.project['teplo'])
	  self.ws.write(self.result, 31, task.project['ohrana'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'VSE42.RU')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 29, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  #print oper
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  #if self.result > 10:
	       #self.stop()

     
bot = Brsn_Zem(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
bot.run()
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')    
command = 'mount -a'# cifs //192.168.1.6/d /home/oleg/pars -o username=oleg,password=1122,_netdev,rw,iocharset=utf8,file_mode=0777,dir_mode=0777' #'mount -a -O _netdev'
p = os.system('echo %s|sudo -S %s' % ('1122', command))
print p
time.sleep(2)
#bot.workbook.close()
workbook.close()
print('Done!') 
time.sleep(5)
os.system("/home/oleg/pars/small/vse_com.py")







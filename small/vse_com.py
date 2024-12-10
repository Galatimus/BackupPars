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
import os
import time
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


workbook = xlsxwriter.Workbook(u'comm/0001-0081_00_C_001-0009_VSE42.xlsx')

    

class Cian_Zem(Spider):
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
	  self.ws.write(0, 7, u"СЕГМЕНТ")
	  self.ws.write(0, 8, u"ТИП_ПОСТРОЙКИ")
	  self.ws.write(0, 9, u"НАЗНАЧЕНИЕ_ОБЪЕКТА")
	  self.ws.write(0, 10, u"ПОТРЕБИТЕЛЬСКИЙ_КЛАСС")
	  self.ws.write(0, 11, u"СТОИМОСТЬ")
	  self.ws.write(0, 12, u"ПЛОЩАДЬ")
	  self.ws.write(0, 13, u"ЭТАЖ")
	  self.ws.write(0, 14, u"ЭТАЖНОСТЬ")
	  self.ws.write(0, 15, u"ГОД_ПОСТРОЙКИ")
	  self.ws.write(0, 16, u"МАТЕРИАЛ_СТЕН")
	  self.ws.write(0, 17, u"ВЫСОТА_ПОТОЛКА")
	  self.ws.write(0, 18, u"СОСТОЯНИЕ")
	  self.ws.write(0, 19, u"БЕЗОПАСНОСТЬ")
	  self.ws.write(0, 20, u"ГАЗОСНАБЖЕНИЕ")
	  self.ws.write(0, 21, u"ВОДОСНАБЖЕНИЕ")
	  self.ws.write(0, 22, u"КАНАЛИЗАЦИЯ")
	  self.ws.write(0, 23, u"ЭЛЕКТРОСНАБЖЕНИЕ")
	  self.ws.write(0, 24, u"ТЕПЛОСНАБЖЕНИЕ")
	  self.ws.write(0, 25, u"ОПИСАНИЕ")
	  self.ws.write(0, 26, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 27, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 28, u"ТЕЛЕФОН")
	  self.ws.write(0, 29, u"КОНТАКТНОЕ_ЛИЦО")
	  self.ws.write(0, 30, u"КОМПАНИЯ")
	  self.ws.write(0, 31, u"ДАТА_РАЗМЕЩЕНИЯ")
	  self.ws.write(0, 32, u"ДАТА_ОБНОВЛЕНИЯ")
	  self.ws.write(0, 33, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 34, u"ОПЕРАЦИЯ")
	  self.ws.write(0, 35, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(1,57):
               yield Task ('post',url='http://dom.vse42.ru/property/commercial/page%d'%x,network_try_count=100)
	       
	       
     def task_post(self,grab,task):
	  for elem in grab.doc.select('//a[@class="obj-adv-url"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True, network_try_count=100)	       
                
     def task_item(self, grab, task):
	  try:
	       #dt = grab.doc.select(u'//div[contains(text(),"Город:")]/span').text()
	       sub = u'Кемеровская Область'#reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
	  except DataNotFound:
	       sub = ''
	  try:
	       ray = grab.doc.select(u'//span[@class="beforeheader"]').text()
	     #print ray 
	  except IndexError:
	       ray = ''          
	  try:
	       #if  grab.doc.select(u'//em/a[2][contains(text(),"р-н")]').exists()==True:
	       punkt= u'Кемерово'#grab.doc.select(u'//div[contains(text(),"Город:")]/span').text()#.split(', ')[2]
	       
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
	  
	  trassa = ''       
	  try:
	       udal = grab.doc.select(u'//div[contains(text(),"Назначение")]/following-sibling::div[1]').text()
	  except DataNotFound:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//div[contains(text(),"Цена")]/following-sibling::div[1]').text()
	  except IndexError:
	       price = ''   
	  try:
	       plosh = grab.doc.select(u'//div[contains(text(),"Площадь объекта")]/following-sibling::div[1]').text()
	  except DataNotFound:
	       plosh = ''
	  
	  try:
	       et = grab.doc.select(u'//label[contains(text(),"Этаж:")]/following-sibling::span').text().split('/')[0]
	  except IndexError:
	       et = ''
	  try:
	       et2 = grab.doc.select(u'//label[contains(text(),"Этаж:")]/following-sibling::span').text().split('/')[1]
	  except IndexError:
	       et2 = ''
	  
	  try:
	       mat = grab.doc.select(u'//label[contains(text(),"Материал стен:")]/following-sibling::span').text()
	  except IndexError:
	       mat = ''
          try:
               godp = grab.doc.select(u'//div[contains(text(),"Состояние ремонта")]/span').text()
          except IndexError:
               godp = ''	       
	       
	       
	  try:
	       ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
	  except DataNotFound:
	       ohrana =''
	  try:
	       gaz = re.sub(u'^.*(?=газ)','', grab.doc.select(u'//*[contains(text(), "газ")]').text())[:3].replace(u'газ',u'есть')
	  except DataNotFound:
	       gaz =''
	  try:
	       voda = re.sub(u'^.*(?=вод)','', grab.doc.select(u'//*[contains(text(), "вод")]').text())[:3].replace(u'вод',u'есть')
	  except DataNotFound:
	       voda =''
	  try:
	       kanal = re.sub(u'^.*(?=санузел)','', grab.doc.select(u'//*[contains(text(), "санузел")]').text())[:7].replace(u'санузел',u'есть')
	  except DataNotFound:
	       kanal =''
	  try:
	       elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
	  except DataNotFound:
	       elek =''
	  try:
	       teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
	  except IndexError:
	       teplo =''
	       
	  try:
	       oper = grab.doc.select(u'//h1[@class="title"]').text().split(' ')[0].replace(u'Сдается',u'Аренда').replace(u'Продается',u'Продажа')  
	  except IndexError:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//h3[contains(text(),"Дополнительная информация")]/following-sibling::div[1]').text()#.replace(u'Описание','')  
	  except IndexError:
	       opis = ''
	       
	
	       
	  try:
	        
	       lico = grab.doc.select(u'//span[@class="ffio"]').text()
	       #except IndexError:
		    #lico = grab.doc.select(u'//td[contains(text(),"Агент:")]/following-sibling::td').text()
	  except IndexError:
	       lico = ''
	       
	  try:
	       comp = grab.doc.select(u'//div[contains(text(),"Опубликовал")]/following-sibling::div[1]').text()#.split(' :: ')[0]
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
		    vid = '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=int(re.sub(u'[^\d]','',dt1))))
	       elif 'день' in dt1:
		    vid = '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=int(re.sub(u'[^\d]','',dt1))))
	       elif 'дня' in dt1:
		    vid = '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=int(re.sub(u'[^\d]','',dt1))))
	       else:
		    vid = datetime.today().strftime('%d.%m.%Y')
	  except DataNotFound:
	       vid = '' 	       
		    
	  
						   
	       
	  projects = {'url': task.url,
                      'sub': sub,
                      'rayon': ray,
                      'punkt': punkt,
                      'teritor': ter,
                      'ulica': uliza,
	              'dom': dom,
                      'trassa': trassa,
                      'udal': udal,
                      'cena': price,
                      'plosh':plosh,
	              'et': et,
	              'ets': et2,
	              'mat': mat,
	              'god':godp,
                      'vid': vid,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'lico':lico,
	              'phone':random.choice(list(open('../phone.txt').read().splitlines())),
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
	  
	  #try:
	       ##json_data = json.loads(grab.response.body)
	       #phone = re.sub('[^\d\+]','',grab.response.json['text'])
	  #except (IndexError,ValueError):
	       #phone=''	       
          
          
	  #yield Task('write',project=task.project,phone=phone,grab=grab)
	  yield Task('write',project=projects,grab=grab)
            
     def task_write(self,grab,task):
	  print('*'*50)
	  print  task.project['sub']
	  print  task.project['punkt']
	  print  task.project['teritor']
	  print  task.project['ulica']
	  print  task.project['dom']
	  print  task.project['trassa']
	  print  task.project['udal']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['et']
	  print  task.project['ets']
	  print  task.project['mat']
	  print  task.project['god']
	  print  task.project['ohrana']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanaliz']
	  print  task.project['electr']
	  print  task.project['teplo']
	  print  task.project['opis']
	  print  task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	  print  task.project['vid']
	  print  task.project['rayon']
	  print  task.project['oper']
	  
	  #global result
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 35, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 3, task.project['teritor'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  #self.ws.write(self.result, 9, task.project['trassa'])
	  self.ws.write(self.result, 9, task.project['udal'])
	  self.ws.write(self.result, 34, task.project['oper'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 13, task.project['et'])
	  self.ws.write(self.result, 14, task.project['ets'])
	  self.ws.write(self.result, 18, task.project['god'])
	  self.ws.write(self.result, 16, task.project['mat'])	  
	  self.ws.write(self.result, 32, task.project['vid'])
	  self.ws.write(self.result, 20, task.project['gaz'])
	  self.ws.write(self.result, 21, task.project['voda'])
	  self.ws.write(self.result, 22, task.project['kanaliz'])
	  self.ws.write(self.result, 23, task.project['electr'])
	  self.ws.write(self.result, 24, task.project['teplo'])
	  self.ws.write(self.result, 19, task.project['ohrana'])	       
	  self.ws.write(self.result, 25, task.project['opis'])
	  self.ws.write(self.result, 26, u'VSE42.RU')
	  self.ws.write_string(self.result, 27, task.project['url'])
	  self.ws.write(self.result, 28, task.project['phone'])
	  self.ws.write(self.result, 29, task.project['lico'])
	  self.ws.write(self.result, 30, task.project['company'])
	  self.ws.write(self.result, 31, task.project['data'])
	  self.ws.write(self.result, 33, datetime.today().strftime('%d.%m.%Y'))
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  #print oper
	  print('*'*50)	       
	  self.result+= 1
	       
	  #if self.result >= 15:
	       #self.stop()	       	       
	 

     
bot = Cian_Zem(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=500, connect_timeout=5000)
bot.run()
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
workbook.close()
print('Done!')
time.sleep(5)
os.system("/home/oleg/pars/small/ya39_zem.py")








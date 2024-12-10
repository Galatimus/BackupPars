#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import math
import re
import os
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'zag/Cian_Загород.xlsx')




class Cian_Zem(Spider):
     def prepare(self):
	  
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
	  self.ws.write(0, 9, u"КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
	  self.ws.write(0, 10, u"ТИП_ОБЪЕКТА")
	  self.ws.write(0, 11, u"ОПЕРАЦИЯ")
	  self.ws.write(0, 12, u"СТОИМОСТЬ")
	  self.ws.write(0, 13, u"ЦЕНА_М2")
	  self.ws.write(0, 14, u"ПЛОЩАДЬ_ДОМА")
	  self.ws.write(0, 15, u"КОЛИЧЕСТВО_КОМНАТ")
	  self.ws.write(0, 16, u"ЭТАЖНОСТЬ")
	  self.ws.write(0, 17, u"МАТЕРИАЛ_СТЕН")
	  self.ws.write(0, 18, u"ГОД_ПОСТРОЙКИ")
	  self.ws.write(0, 19, u"ПЛОЩАДЬ_УЧАСТКА")
	  self.ws.write(0, 20, u"ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	  self.ws.write(0, 21, u"ГАЗОСНАБЖЕНИЕ")
	  self.ws.write(0, 22, u"ВОДОСНАБЖЕНИЕ")
	  self.ws.write(0, 23, u"КАНАЛИЗАЦИЯ")
	  self.ws.write(0, 24, u"ЭЛЕКТРОСНАБЖЕНИЕ")
	  self.ws.write(0, 25, u"ТЕПЛОСНАБЖЕНИЕ")
	  self.ws.write(0, 26, u"ЛЕС")
	  self.ws.write(0, 27, u"ВОДОЕМ")
	  self.ws.write(0, 28, u"БЕЗОПАСНОСТЬ")
	  self.ws.write(0, 29, u"ОПИСАНИЕ")
	  self.ws.write(0, 30, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 31, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 32, u"ТЕЛЕФОН")
	  self.ws.write(0, 33, u"КОНТАКТНОЕ_ЛИЦО")
	  self.ws.write(0, 34, u"КОМПАНИЯ")
	  self.ws.write(0, 35, u"ДАТА_РАЗМЕЩЕНИЯ")
	  self.ws.write(0, 36, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 37, u"МЕСТОПОЛОЖЕНИЕ")    
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  l= open('cian_zag.txt').read().splitlines()
	  self.dc = len(l)
	  print self.dc
	  for line in l:
	       yield Task ('item',url=line,refresh_cache=True,network_try_count=100)

     def task_item(self, grab, task):
	  
	  try:
	       sub = grab.doc.select(u'//address[contains(@class,"address")]').text().split(', ')[0]
	  except IndexError:
	       sub = ''	       

	  try:
	       ray = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"р-н ")]').text()
	     #print ray 
	  except DataNotFound:
	       ray = ''          
	  try:
	       if sub == u'Москва':
		    punkt= u'Москва'
	       elif sub == u'Санкт-Петербург':
	            punkt= u'Санкт-Петербург'
	       elif sub == u'Севастополь':
	            punkt= u'Севастополь'
	       else:	       
		    if  grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"р-н ")]').exists()==True:
			 punkt= grab.doc.select(u'//address[contains(@class,"address")]').text().split(', ')[2]
		    else:
			 punkt= grab.doc.select(u'//address[contains(@class,"address")]').text().split(', ')[1]
	  except IndexError:
	       punkt = ''
	       
	  try:
	       #if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text()," район")]').exists()==True:
	       ter= grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"р-н ")]').text()#.split(', ')[3].replace(u'улица','')
	       #else:
		    #ter= ''#grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[2]
	  except IndexError:
	       ter =''
	       
	  try:
	       try:
		    try:
			 try:
			      try:
				   try:
					try:
					     uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"ул.")]').text()
					except IndexError:
					     uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"пер.")]').text()
				   except IndexError:
					uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"просп.")]').text()
			      except IndexError:
				   uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"б-р")]').text()
			 except IndexError:
			      uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"бул.")]').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"проезд")]').text()
	       except IndexError:
		    uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"наб.")]').text()
	  except IndexError:
	       uliza = ''
	  try:
	       dom = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(@href,"house")]').text()
	  except DataNotFound:
	       dom = ''
	       
	  try:
	       trassa = grab.doc.select(u'//h1').text()
	       if 'дом' in trassa:
		    trassa = 'Дом'
	       else:
		    trassa = 'Коттедж' 
	  except DataNotFound:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//div[contains(text(),"Общая")]/following-sibling::div').text()
	  except DataNotFound:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//span[@itemprop="price"]').text()
	  except DataNotFound:
	       price = ''
	       
	  try:
	       try:
		    plosh = grab.doc.select(u'//span[contains(text(),"Площадь участка")]/following-sibling::span[1]').text()#.replace(u'участок ','')
	       except IndexError:
		    plosh = grab.doc.select(u'//div[contains(text(),"Участок")]/following-sibling::div').text()
	  except IndexError:
	       plosh = ''
	       
	  try:
	       vid = grab.doc.select(u'//span[contains(text(),"Этажей в доме")]/following-sibling::span[1]').text()
	  except DataNotFound:
	       vid = '' 
	       
	       
	  try:
	       ohrana = grab.doc.select(u'//span[contains(text(),"Тип дома")]/following-sibling::span[1]').text()
	  except DataNotFound:
	       ohrana =''
	  try:
	       gaz = grab.doc.select(u'//div[contains(text(),"Построен")]/following-sibling::div').text()
	  except DataNotFound:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"ш.")]').text()
	  except DataNotFound:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//th[contains(text(),"Канализация:")]/following-sibling::td').text()
	  except DataNotFound:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//th[contains(text(),"Электричество:")]/following-sibling::td').text()
	  except DataNotFound:
	       elek =''
	  try:
	       teplo = grab.doc.select(u'//address[contains(@class,"address")]').text().replace(u'На карте','')
	  except DataNotFound:
	       teplo =''
	       
	  try:
	       if 'sale' in task.url:
		    oper = u'Продажа' 
	       elif 'rent' in task.url:
		    oper = u'Аренда'     
	  except DataNotFound:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//p[@itemprop="description"]').text()#.split(u'Показать телефон')[0] 
	  except IndexError:
	       opis = ''
	       
	  try:
	       phone = grab.doc.rex_text(u'href="tel:(.*?)"') 
	  except DataNotFound:
	       phone = ''
	       
	  try:
	       try:
		    lico = grab.doc.select(u'//h3[@class="realtor-card__title"][contains(text(),"Представитель: ")]').text().replace(u'Представитель: ','')
	       except IndexError:
		    lico = grab.doc.select(u'//h2').text() 
	  except IndexError:
	       lico = ''
	       
	  try:
	       try:
		    comp = grab.doc.select(u'//a[contains(@href,"company")]/h2').text()
	       except IndexError:
		    comp = grab.doc.select(u'//h4[@class="realtor-card__subtitle"]').text()
	  except IndexError:
	       comp = ''
	       
	  try:
	       data = re.sub(u'[^\d\-]','',grab.doc.rex_text(u'editDate(.*?)T')).replace('-','.')
	    #print data
	  except DataNotFound:
	       data = ''
		    
	  
						   
	       
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
                      'vid': vid,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'phone':phone.replace(u'79311111111',''),
                      'lico':lico,
                      'company':comp,
                      'data':data,
                      'oper':oper
                      }
	  
	  yield Task('write',project=projects,grab=grab)
	    
     def task_write(self,grab,task):
	  if task.project['sub'] <> '': 
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
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['teplo']
	       #print  task.project['oper']
	       
	       #global result
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 10, task.project['trassa'])
	       self.ws.write(self.result, 14, task.project['udal'])
	       self.ws.write(self.result, 11, task.project['oper'])
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 19, task.project['plosh'])
	       self.ws.write(self.result, 16, task.project['vid'])
	       self.ws.write(self.result, 18, task.project['gaz'])
	       self.ws.write(self.result, 7, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 18, task.project['electr'])
	       self.ws.write(self.result, 37, task.project['teplo'])
	       self.ws.write(self.result, 17, task.project['ohrana'])	       
	       self.ws.write(self.result, 29, task.project['opis'])
	       self.ws.write(self.result, 30, u'ЦИАН')
	       self.ws.write_string(self.result, 31, task.project['url'])
	       self.ws.write(self.result, 32, task.project['phone'])
	       self.ws.write(self.result, 33, task.project['lico'])
	       self.ws.write(self.result, 34, task.project['company'])
	       self.ws.write(self.result, 35, task.project['data'])
	       self.ws.write(self.result, 36, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       #print task.sub
	       
	       print 'Ready - '+str(self.result)+'/'+str(self.dc)
	       print 'Tasks - %s' % self.task_queue.size()
	       print  task.project['oper']
	       print('*'*50)	       
	       self.result+= 1
		    
	       #if self.result > 20:
		    #self.stop()

     
bot = Cian_Zem(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
try:
     bot.run()
except KeyboardInterrupt:
     pass
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
command = 'mount -a'
os.system('echo %s|sudo -S %s' % ('1122', command))
time.sleep(2)
workbook.close()
print('Done')

   








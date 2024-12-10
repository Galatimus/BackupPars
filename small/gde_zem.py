#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import random
import re
import os
from sub import conv
import time
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'zemm/0001-0002_00_У_001-0120_GDE-RU.xlsx')


class Farpost_Zem(Spider):
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
	  self.ws.write(0, 29, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 30, u"КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
	  self.ws.write(0, 31, u"ЗАГОЛОВОК")
	  self.ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
	  #self.r = conv     
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(1,300):#78
	       yield Task ('post',url='https://gde.ru/nedvizhimost/uchastki/prodam?page=%d'%x,refresh_cache=True,network_try_count=50)
	  for x1 in range(1,12):#4
	       yield Task ('post',url='https://gde.ru/nedvizhimost/uchastki/sdam?page=%d'%x1,refresh_cache=True,network_try_count=50)
	       
     def task_post(self,grab,task):    
	  for elem in grab.doc.select(u'//div[@class="title"]/a'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       #if 'uchastok' in ur:
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=50)
	  #yield Task("page", grab=grab,refresh_cache=True,network_try_count=10)
	 
	 
     def task_page(self,grab,task):
	  try:
	       pg = grab.doc.select(u'//li[@class="text control forward last-none"]/a[contains(text(),"вперёд")]')
	       u = grab.make_url_absolute(pg.attr('href'))
	       yield Task ('post',url= u,refresh_cache=True,network_try_count=10)
	  except IndexError:
	       print('*'*100)
	       print '!!!','NO PAGE NEXT','!!!'
	       print('*'*100)
        
     def task_item(self, grab, task):

	  try:
	       r = grab.doc.select(u'//div[@class="popButton cityPop"]').text()
	       if r.find(u'район')>=0:
		    ray = r
	       else:
		    ray=''
	     #print ray 
	  except DataNotFound:
	       ray = ''          
	  try:
	       punkt= grab.doc.select(u'//a[contains(@href,"city")]').text()
	  except IndexError:
	       punkt = ''
	       
	  try:
	       ter= grab.doc.select(u'//div[contains(text(),"Район")]/following-sibling::div/span').text()
	  except IndexError:
	       ter =''
	       
	  try:
	       
	       uliza = grab.doc.select(u'//dt[contains(text(),"Категория земель")]/following-sibling::dd').text().split(' (')[1].replace(')','')
	       #else:
		    #uliza = ''
	  except IndexError:
	       uliza = ''
	       
          try:
               dom = grab.doc.select(u'//a[@itemprop="item"]/span[contains(text(),"участок")]').text().split(' ')[0].replace(u'Продам',u'Продажа').replace(u'Сдам',u'Аренда')
          except IndexError:
               dom = ''
	       
	  try:
	       trassa = grab.doc.select(u'//title').text()
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//dt[contains(text(),"Расстояние до города")]/following-sibling::dd').text()
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//span[@class="price"]').text()#.replace(u'a',u'р.')
	  except IndexError:
	       price = ''
	       
	  
	       
	  try:
	       plosh = grab.doc.select(u'//dt[contains(text(),"Площадь")]/following-sibling::dd').text()
	  except IndexError:
	       plosh = ''
	       
	  
	  
	  
	       
	  try:
	       vid = grab.doc.select(u'//dt[contains(text(),"Категория земель")]/following-sibling::dd').text().split(' (')[0].replace(')','')
	  except IndexError:
	       vid = '' 
	       
	       
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
	       teplo = grab.doc.select(u'//div[@class="address"]').text()
	  except DataNotFound:
	       teplo =''
	       
	                
	      
		    
	  try:
	       opis = grab.doc.select(u'//p[@itemprop="description"]').text() 
	  except IndexError:
	       opis = ''
	       
	  
	       
	  try:
	       lico = grab.doc.select(u'//a[@class="user"]').text()
	  except IndexError:
	       lico = ''
	       
	  try:
	       comp = grab.doc.select(u'//span[@class="inplace"][contains(@data-field,"realtyStatus")]').text()
	  except IndexError:
	       comp = ''
	       
	  try:
	       data = grab.doc.select(u'//div[@class="date"]').text().replace(u'с ','')
	    #print data
	  except IndexError:
	       data = ''
		    
	  
	  
	       
	  projects = {'url': task.url,
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
	              'phone': random.choice(list(open('../phone.txt'))),
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'lico':lico,
                      'company':comp,
                      'data':data }
	  
	  
	  try:
	       link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode=Россия, '+punkt
	       yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=50)
	  except IndexError:
	       pass   
	  
	  
     def task_adres(self, grab, task):
	  try:   
	       sub = grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["Address"]["Components"][2]["name"]
	  except (ValueError,IndexError,TypeError,KeyError,AttributeError):
	       sub = ''

          
	  yield Task('write',project=task.project,sub=sub,grab=grab)
            
     def task_write(self,grab,task):
	  print('*'*50)
	  print  task.sub
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['teritor']
	  print  task.project['ulica']
	 
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

	  
	  #global result
	  self.ws.write(self.result, 0, task.sub)
	  self.ws.write(self.result, 1, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 3, task.project['teritor'])
	  self.ws.write(self.result, 14, task.project['ulica'])
	  self.ws.write(self.result, 9, task.project['dom'])
	  self.ws.write(self.result, 31, task.project['trassa'])
	  self.ws.write(self.result, 8, task.project['udal'])
	  #self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write_string(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 13, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 32, task.project['teplo'])
	  self.ws.write(self.result, 20, task.project['ohrana'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'GDE.RU')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25,task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  #self.ws.write(self.result, 31, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  print  task.project['dom']
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  if self.result > 8000:
	       self.stop()

     
bot = Farpost_Zem(thread_number=15,network_try_limit=500)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=500, connect_timeout=500)
try:
     bot.run()
except KeyboardInterrupt:
     pass
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
workbook.close()
print('Done')

time.sleep(5)
os.system("/home/oleg/pars/small/adv_zem.py")






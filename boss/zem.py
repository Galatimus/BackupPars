#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
import random
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import time
import os
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)









class Nedvizhka_Zem(Spider):
     def prepare(self):
	  self.workbook = xlsxwriter.Workbook(u'0001-0002_00_У_001-0206_BEBOSS.xlsx')
	  self.ws = self.workbook.add_worksheet()
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
	  self.ws.write(0, 21, u"КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
	  self.ws.write(0, 22, u"ОПИСАНИЕ")
	  self.ws.write(0, 23, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 24, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 25, u"ТЕЛЕФОН")
	  self.ws.write(0, 26, u"КОНТАКТНОЕ_ЛИЦО")
	  self.ws.write(0, 27, u"КОМПАНИЯ")
	  self.ws.write(0, 28, u"ДАТА_РАЗМЕЩЕНИЯ")
	  self.ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
	  self.ws.write(0, 30, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 31, u"ШИРОТА_ИСХ")
	  self.ws.write(0, 32, u"ДОЛГОТА_ИСХ")	  
	  self.ws.write(0, 33, u"МЕСТОПОЛОЖЕНИЕ")
	  self.result= 1
       
       
       
	 

     def task_generator(self):
	  for x in range(1,11):
	       yield Task ('post',url='https://www.beboss.ru/kn/ru_land_rent?page=%d'%x,refresh_cache=True,network_try_count=100)
	  for z in range(1,63):
	       yield Task ('post',url='https://www.beboss.ru/kn/ru_land_sell?page=%d'%z,refresh_cache=True,network_try_count=100)	  

	     
	  
       
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//div[@class="obj__right"]/a[contains(text(),"Подробнее")]'):
	       ur = grab.make_url_absolute(elem.attr('href'))
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True, network_try_count=100)


     def task_item(self, grab, task):
	  try:
	       ray = grab.doc.select(u'//p[@class="object-addresses"]').text()
	  except IndexError:
	       ray = ''          

	  try:
	       metro = grab.doc.select(u'//dt[contains(text(),"Район")]/following-sibling::dd').text()
	    #print rayon
	  except IndexError:
	       metro = ''
	      
	  try:
	       metro_min = grab.doc.select(u'//ul[@class="breadcrumbs breadcrumbs_dealer"]/li[2]/a').text().split(' ')[0]
	  except IndexError:
	       metro_min = ''
	      
	  try:
	       metro_tr = grab.doc.select(u'//p[@class="kn-obj-title b-franchise-hide-mobile"]').text().split(': ')[1]
	  except IndexError:
	       metro_tr = ''

	  try:
               price = grab.doc.select(u'//p[@class="kn-obj-info b-franchise-hide-mobile"]/text()[1]').text().split(': ')[1]
	  except IndexError:
	       price = ''
	  try:
	       plosh_ob = grab.doc.select(u'//p[@class="kn-obj-info b-franchise-hide-mobile"]/text()[2]').text().split(': ')[1]
	  except IndexError:
	       plosh_ob = ''
	  try:
	       et = grab.doc.select(u'//p[@itemprop="description"]').text()
	  except IndexError:
	       et = '' 
	      
	  try:
	       etagn = grab.doc.select(u'//p[@class="franchise-person__name"]').text()
	  except IndexError:
	       etagn = ''
	  try:
	       opis = grab.doc.select(u'//p[@class="kn-company-short__text"]/a[contains(@href, "company")]').text()
	  except IndexError:
	       opis = ''
	  try:
	       lico = grab.doc.select(u'//span[@class="kn-type-object__date"]').text().split(', ')[0].replace(u'Обновлено ','')
	  except IndexError:
	       lico = ''
	       
          comp = ''
	  data = ''

	  try:
	       lat = grab.doc.rex_text(u'id="lat" value="(.*?)"')
	  except IndexError:
	       lat =''
	       
	  try:
	       lng = grab.doc.rex_text(u'id="lng" value="(.*?)"')
	  except IndexError:
	       lng =''		    

          
	  
	  projects = {'rayon': ray,
                      'metro': metro,
                      'naz': metro_min,		           
                      'tran': metro_tr,
                      'cena': price,		           
                      'plosh_ob':plosh_ob,		           
                      'etach': et,
                      'etashost': etagn,      
                      'opis':opis,
                      'url':task.url,
                      'phone': random.choice(list(open('../phone.txt').read().splitlines())),
                      'dol': lat,
                      'shir': lng,	                   
                      'lico':lico,
                      'company':comp,
                      'data':data}
	  try:
	       link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ray
	       yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=100)
	  except IndexError:
	       yield Task('adres',grab=grab,project=projects)
	
	
     def task_adres(self, grab, task):
     
	  try:   
	       sub= grab.doc.rex_text(u'AdministrativeAreaName":"(.*?)"')
	  except IndexError:
	       sub = ''	    
	  try:   
	       punkt= grab.doc.rex_text(u'LocalityName":"(.*?)"')
	  except IndexError:
	       punkt = ''
	  try:
	       ter=  grab.doc.rex_text(u'DependentLocalityName":"(.*?)"')
	  except IndexError:
	       ter =''
	  try:
	       uliza=grab.doc.rex_text(u'ThoroughfareName":"(.*?)"')
	  except IndexError:
	       uliza = ''
	  try:
	       dom = grab.doc.rex_text(u'PremiseNumber":"(.*?)"')
	  except IndexError:
	       dom = ''
     
	  project2 ={'punkt':punkt,
	             'sub': sub,
	             'teritor': ter,
	             'ulica':uliza,
	             'dom':dom}
	  
	  
	  
	  
	  yield Task('write',project=task.project,proj=project2,grab=grab)

     def task_write(self,grab,task):
	  
	  print('*'*50)	       
	  print  task.proj['sub']
          print  task.proj['punkt']
          print  task.proj['teritor']
          print  task.proj['ulica']	    
          print  task.proj['dom']
	  print  task.project['metro']
	  print  task.project['naz']	      
	  print  task.project['tran']
	  print  task.project['cena']	       
	  print  task.project['plosh_ob']	       
	  print  task.project['etach']
	  print  task.project['etashost']
	  print  task.project['opis']
	  print  task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['rayon']
	  
    
	  self.ws.write(self.result, 0,task.proj['sub'])
	  self.ws.write(self.result, 1,task.project['metro'])
	  self.ws.write(self.result, 2,task.proj['punkt'])
	  self.ws.write(self.result, 3,task.proj['teritor'])
	  self.ws.write(self.result, 4,task.proj['ulica'])
	  self.ws.write(self.result, 5,task.proj['dom'])	 
	  self.ws.write(self.result, 9,task.project['naz'])
	  self.ws.write(self.result, 12,task.project['tran'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 11, task.project['plosh_ob'])
	  self.ws.write(self.result, 33, task.project['rayon'])
	  self.ws.write(self.result, 31, task.project['dol'])
	  self.ws.write(self.result, 32, task.project['shir'])
	  self.ws.write(self.result, 22, task.project['etach'])
	  self.ws.write(self.result, 26, task.project['etashost'])
	  self.ws.write(self.result, 27, task.project['opis'])
	  self.ws.write(self.result, 23, u'БИБОСС')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 29, task.project['lico'])
	  #self.ws.write(self.result, 9, task.project['company'])
	  #self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	 
	  
	  print('*'*50)
	  print 'Ready - '+str(self.result)
	  print 'Tasks - %s' % self.task_queue.size()
	  print('*'*50)
	  self.result+= 1


bot = Nedvizhka_Zem(thread_number=5,network_try_limit=1000)
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
bot.workbook.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/boss/com.py")





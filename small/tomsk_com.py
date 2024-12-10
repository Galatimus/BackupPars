#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import random
import re
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


workbook = xlsxwriter.Workbook(u'comm/0001-0002_00_C_001-0295_TOMSK9.xlsx')

    

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
	  self.ws.write(0, 12, u"ИЗМЕНЕНИЕ_СТОИМОСТИ")
	  self.ws.write(0, 13, u"ДОПОЛНИТЕЛЬНЫЕ_КОММЕРЧЕСКИЕ_УСЛОВИЯ")
	  self.ws.write(0, 14, u"ПЛОЩАДЬ")
	  self.ws.write(0, 15, u"ЭТАЖ")
	  self.ws.write(0, 16, u"ЭТАЖНОСТЬ")
	  self.ws.write(0, 17, u"МАТЕРИАЛ_СТЕН")
	  self.ws.write(0, 18, u"ОПИСАНИЕ")
	  self.ws.write(0, 19, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 20, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 21, u"ТЕЛЕФОН")
	  self.ws.write(0, 22, u"КОНТАКТНОЕ_ЛИЦО")
	  self.ws.write(0, 23, u"КОМПАНИЯ")
	  self.ws.write(0, 24, u"МЕСТОПОЛОЖЕНИЕ")
	  self.ws.write(0, 25, u"МЕСТОРАСПОЛОЖЕНИЕ")
	  self.ws.write(0, 26, u"БЛИЖАЙШАЯ_СТАНЦИЯ_МЕТРО")
	  self.ws.write(0, 27, u"РАССТОЯНИЕ_ДО_БЛИЖАЙШЕЙ_СТАНЦИИ_МЕТРО")
	  self.ws.write(0, 28, u"ОПЕРАЦИЯ")
	  self.ws.write(0, 29, u"ДАТА_РАЗМЕЩЕНИЯ")
	  self.ws.write(0, 30, u"ДАТА_ОБНОВЛЕНИЯ")
	  self.ws.write(0, 31, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 32, u"ЦЕНА_ЗА_М2")
	  self.ws.write(0, 33, u"ЗАГОЛОВОК")
	  self.ws.write(0, 34, u"ШИРОТА_ИСХ")
	  self.ws.write(0, 35, u"ДОЛГОТА_ИСХ")
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(1,101):
               yield Task ('post',url='https://www.tomsk.ru09.ru/realty/?type=1&otype=3&page=%d'%x,refresh_cache=True,network_try_count=100)
          for x1 in range(1,102):
	       yield Task ('post',url='https://www.tomsk.ru09.ru/realty/?type=2&otype=3&page=%d'%x1,refresh_cache=True,network_try_count=100)
         
                 
        
        
            
            
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//a[@class="visited_ads"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	 
        
        
     def task_item(self, grab, task):
	  try:
	       sub = u'Томская область'
	  except DataNotFound:
	       sub = ''
	  try:
	       ray = grab.doc.select(u'//td[@class="-padding-left"]/span[contains(text()," район")]').text()
	  except IndexError:
	       ray = ''          
	  try:
	       punkt= u'Томск'
	  except IndexError:
	       punkt = ''
	       
	  try:
	       ter= grab.doc.select(u'//span[contains(text(),"Вид")]/following::td[1]').text()
	  except IndexError:
	       ter =''
	       
	  try:
	       uliza = grab.doc.select(u'//td[@class="-padding-left"]/span/a').text()
	  except IndexError:
	       uliza = ''
	  try:
	       dom = grab.doc.select(u'//div[@class="realty_detail_price inline"]/following::td[1]/text()[2]').text().replace(').','')
	  except IndexError:
	       dom = ''       
	  try:
	       trassa = grab.doc.select(u'//span[contains(text(),"Тип объекта")]/following::td[1]').text()
	  except IndexError:
	       trassa = ''       
	  try:
	       udal = grab.doc.select(u'//td[@class="tableh1"]/h3').text()
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//td[@class="realty_detail_price"]/div').text()
	  except IndexError:
	       price = ''   
	  try:
	       plosh = grab.doc.select(u'//span[contains(text(),"Общая площадь")]/following::td[1]').text()
	  except IndexError:
	       plosh = ''
	  try:
	       vid = grab.doc.select(u'//div[@id="catalog-card__map"]').attr('data-coord').split('x')[0]
	  except IndexError:
	       vid = '' 
	  try:
	       et = grab.doc.select(u'//span[contains(text(),"Этаж/этажность")]/following::td[1]').text().split('/')[0]
	  except IndexError:
	       et = ''
	  try:
	       et2 = grab.doc.select(u'//span[contains(text(),"Этажность")]/following::td[1]').number()
	  except IndexError:
	       et2 = ''
	  
	  try:
	       mat = grab.doc.select(u'//td[contains(text(),"Есть")]').text()
	  except IndexError:
	       mat = ''
          try:
               godp = grab.doc.select(u'//span[contains(text(),"Цена за квадратный метр")]/following::dd[1]').text()
          except IndexError:
               godp = ''	       
	  try:
	       if 'Продам' in udal:
		    oper = u'Продажа' 
	       elif 'Сдам' in udal:
		    oper = u'Аренда'
	       else:
		    oper = ''
	  except IndexError:
	       oper = ''               
	  try:
	       opis = grab.doc.select(u'//strong[contains(text(),"Описание")]/following::tr[2]/td').text() 
	  except IndexError:
	       opis = ''

	  try:
	       lico = grab.doc.select(u'//td[contains(text(),"Контактное")]/following-sibling::td').text()
	  except IndexError:
	       lico = ''
	  try:
	       comp = grab.doc.select(u'//a[contains(@href, "agency")][contains(text(),"гентство")]').text()
	  except IndexError:
	       comp = ''
	  try:
	       data= grab.doc.select(u'//td[@class="realty_detail_date nobr"]').attr('title').split(' ')[0]
	  except IndexError:
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
	              'et': et,
	              'ets': et2,
	              'mat': mat,
	              'god':godp,
                      'vid': vid,
                      'opis':opis+' '+mat,
                      'phone':random.choice(list(open('../phone.txt').read().splitlines())),
                      'lico':lico,
                      'company':comp,
                      'data':data,
                      'oper':oper
                      }
          
	  yield Task('write',project=projects,grab=grab,refresh_cache=True)
            
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
	  print  task.project['et']
	  print  task.project['ets']
	  print  task.project['mat']
	  print  task.project['god']
	  print  task.project['vid']
	  print  task.project['opis']
	  print task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	  print  task.project['oper']
	  
	  #global result
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 3, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 9, task.project['teritor'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 13, task.project['dom'])
	  self.ws.write(self.result, 7, task.project['trassa'])
	  self.ws.write(self.result, 33, task.project['udal'])
	  self.ws.write(self.result, 28, task.project['oper'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  self.ws.write(self.result, 15, task.project['et'])
	  self.ws.write(self.result, 16, task.project['ets'])
	  self.ws.write(self.result, 32, task.project['god'])
	  #self.ws.write(self.result, 17, task.project['mat'])	  
	  self.ws.write(self.result, 34, task.project['vid'])
	  self.ws.write(self.result, 18, task.project['opis'])
	  self.ws.write(self.result, 19, u'Tomsk.Ru09')
	  self.ws.write_string(self.result, 20, task.project['url'])
	  self.ws.write(self.result, 21, task.project['phone'])
	  self.ws.write(self.result, 22, task.project['lico'])
	  self.ws.write(self.result, 23, task.project['company'])
	  self.ws.write(self.result, 29, task.project['data'])
	  self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	   
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  #print oper
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  #if self.result >= 50:
	       #self.stop()
	 

     
bot = Cian_Zem(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
try:
     bot.run()
except KeyboardInterrupt:
     pass
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
workbook.close()
print('Done')
os.system("/home/oleg/pars/small/infoline_com.py")








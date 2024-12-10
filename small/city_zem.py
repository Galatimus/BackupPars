#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
from sub import conv
import time
import os
import math
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logging.basicConfig(level=logging.DEBUG)






oper = u'Продажа'


class Mag_Zem(Spider):
     def prepare(self):
	  self.workbook = xlsxwriter.Workbook(u'zemm/0001-0002_00_У_001-0033_CITYST.xlsx')
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
	  self.ws.write(0, 21, u"ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	  self.ws.write(0, 22, u"ОПИСАНИЕ")
	  self.ws.write(0, 23, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 24, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 25, u"ТЕЛЕФОН")
	  self.ws.write(0, 26, u"КОНТАКТНОЕ_ЛИЦО")
	  self.ws.write(0, 27, u"КОМПАНИЯ")
	  self.ws.write(0, 28, u"ДАТА_РАЗМЕЩЕНИЯ")
	  self.ws.write(0, 29, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 30, u"МЕСТОПОЛОЖЕНИЕ")
	  self.result= 1
       
       
       
	 

     def task_generator(self):
	  yield Task ('next',url='http://magnitogorsk-citystar.ru/change-city',refresh_cache=True,network_try_count=100)
	     
	  
     def task_next(self,grab,task):
	  for el in grab.doc.select(u'//h1/following-sibling::ul/li/a'):
	       urr = grab.make_url_absolute(el.attr('href'))  
	       #print urr
	       yield Task('post', url=urr+'/realty/prodazha-zemelnix-uchastkov/',refresh_cache=True,network_try_count=100)
	       

       
     def task_post(self,grab,task):
	  links = grab.doc.select(u'//a[@class="detail-link"]')
	  for elem in links:
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
          yield Task('page', grab=grab,refresh_cache=True,network_try_count=100)   
	       
     def task_page(self,grab,task):
	  try:
	       pg = grab.doc.select(u'//a[@class="pager__navigation common-link-visited"][contains(text(),"следующая")]')
	       u = grab.make_url_absolute(pg.attr('href'))
	       yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	  except IndexError:
	       print 'no_page'

     def task_item(self, grab, task):
	  try:
	       ray = grab.doc.select(u'//td[contains(text(),"Населенный пункт")]/following-sibling::td').text()
	  except IndexError:
	       ray = ''          
	  try:
	       punkt= grab.doc.select(u'//div[@class="cur-city-name"]').text().title()
	  except IndexError:
	       punkt = ''
	  #try:
	       #ter= grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Микрорайон")]/following-sibling::dd').text()
	  #except IndexError:
	       #ter =''
	  try:
	       uliza = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text().split(', ')[0]
	  except IndexError:
	       uliza = ''
	  try:
	       dom = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text().split(', ')[1]
	  except IndexError:
	       dom = ''
		
	  try:
	       orentir = grab.doc.select(u'//td[contains(text(),"Удаленность от города")]/following-sibling::td').text()+' км'
	  except IndexError:
	       orentir = ''              
	    
	  try:
	       metro = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text()
	    #print rayon
	  except IndexError:
	       metro = ''
	      
	  try:
	       metro_min = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Удаленность от города")]/preceding-sibling::div').text()
	    #print rayon
	  except IndexError:
	       metro_min = ''
	      
	  try:
	       metro_tr = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Категория земли")]/preceding-sibling::div').text()
	  except IndexError:
	       metro_tr = ''

	  try:
	       price = grab.doc.select(u'//td[contains(text(),"Цена")]/following-sibling::td/span').text()+u' р.'
	    #print price + u' руб'	    
	  except IndexError:
	       price = ''
   
	  try:
	       plosh_ob = grab.doc.select(u'//td[contains(text(),"Площадь")]/following-sibling::td').text()
	  except IndexError:
	       plosh_ob = ''

	  try:
	       et = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Назначение земли")]/preceding-sibling::div').text()
	    #print price + u' руб'	    
	  except IndexError:
	       et = '' 
	      
	  try:
	       etagn = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Газ")]/preceding-sibling::div').text()
	    #print price + u' руб'	    
	  except IndexError:
	       etagn = ''
	  try:
	       opis = grab.doc.select(u'//td[@class="note"]').text() 
	  except IndexError:
	       opis = ''
	   
	  try:
	       phone = re.sub('[^\d\,]','',grab.doc.select(u'//span[contains(text(),"тел.:")]/following-sibling::text()').text())
	  except IndexError:
	       phone = ''
	      
	  try:
	       lico = grab.doc.select(u'//div[@class="phone"]/preceding-sibling::div[@class="name"]').text()
	  except IndexError:
	       lico = ''
	       
	  try:
	       comp = grab.doc.select(u'//span[contains(text(),"Агентство недвижимости:")]/following-sibling::div[1]').text()
	  except IndexError:
	       comp = ''
	       
	  try: 
	       con = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
	            (u' мая ',u'.05.'),(u' июня ',u'.06.'),
	            (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
	            (u' января ',u'.01.'),(u' декабря ',u'.12.'),
	            (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
	            (u' февраля ',u'.02.'),(u' октября ',u'.10.'),
	            (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
	            (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]
	       d = grab.doc.select(u'//div[@class="date"]').text().replace(u'Дата подачи: ','').split(u'г.')[0]
	       data = reduce(lambda d, r: d.replace(r[0], r[1]), con, d)
	  except IndexError:
	       data=''

          sub = reduce(lambda punkt, r: punkt.replace(r[0], r[1]), conv, punkt)
      
	  projects = {'sub': sub,
                      'rayon': ray.replace(punkt,'').replace(u'г.',''),
                      'punkt': punkt,
                      'teritor':orentir,
                      'ulica': uliza,
                      'dom': dom,
                      'metro': metro,
                      'naz': metro_min,		           
                      'tran': metro_tr,
                      'cena': price,		           
                      'plosh_ob':plosh_ob,		           
                      'etach': et,
                      'etashost': etagn,      
                      'opis':opis,
                      'url':task.url,
                      'phone':phone,
                      'lico':lico,
                      'company':comp,
                      'data':data}
	
	
	
	  yield Task('write',project=projects,grab=grab)
	
	
	
	
	
	
     def task_write(self,grab,task):
	  
	  print('*'*50)	       
	  print  task.project['sub']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['teritor']
	  print  task.project['ulica']
	  print  task.project['dom']
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
	  print  task.project['data']
	  
    
	  self.ws.write(self.result, 0,task.project['sub'])
	  self.ws.write(self.result, 3,task.project['rayon'])
	  self.ws.write(self.result, 2,task.project['punkt'])
	  self.ws.write(self.result, 8,task.project['teritor'])
	  self.ws.write(self.result, 4,task.project['ulica'])
	  self.ws.write(self.result, 5,task.project['dom'])
	  #self.ws.write(self.result, 7,task.project['metro'])
	  self.ws.write(self.result, 8,task.project['naz'])
	  self.ws.write(self.result, 13,task.project['tran'])
	  self.ws.write(self.result, 9,oper)
	  self.ws.write(self.result, 10, task.project['cena'])
	  #self.ws.write(self.result, 13, task.project['cena_m'])
	  #self.ws.write(self.result, 14, task.project['col_komnat'])
	  self.ws.write(self.result, 12, task.project['plosh_ob'])
	  self.ws.write(self.result, 14, task.project['etach'])
	  self.ws.write(self.result, 15, task.project['etashost'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'СИТИСТАР')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 30, task.project['sub']+
                        ', '+task.project['punkt']+
                        ', '+task.project['rayon']+
                        ', '+task.project['metro'])	       
	 
	  
	  print('*'*10)
	  print 'Ready - '+str(self.result)
	  print 'Tasks - %s' % self.task_queue.size()
	  print('*'*10)
	  self.result+= 1
	  
	  
	  
	  
	  #if self.result > 10:
	       #self.stop()


bot = Mag_Zem(thread_number=5,network_try_limit=2000)
#bot.setup_queue(backend='mongo', database='city',host='192.168.10.200')
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
bot.run()
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
bot.workbook.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/small/city_com.py")

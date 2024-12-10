#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
import base64
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
	  self.workbook = xlsxwriter.Workbook(u'0001-0002_00_У_001-0041_TVOYAD.xlsx')
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
	  self.conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
		    (u' мая ',u'.05.'),(u' июня ',u'.06.'),
		    (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
		    (u' января ',u'.01.'),(u' декабря ',u'.12.'),
		    (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
		    (u' февраля ',u'.02.'),(u' октября ',u'.10.'),
		    (u'25.10.2010', (datetime.today().strftime('%d.%m.%Y'))),
		    (u'1.01.1970','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]	  
	  
	  self.result= 1
       
       
       
	 

     def task_generator(self):
	  for x in range(1000):
	       yield Task ('post',url='https://tvoyadres.ru/nedvizhimost/zemelnye-uchastki/prodazha/?page=%d'%x,refresh_cache=True,network_try_count=100)
	  for z in range(365):
	       yield Task ('post',url='https://tvoyadres.ru/nedvizhimost/zemelnye-uchastki/sdacha/?page=%d'%z,refresh_cache=True,network_try_count=100)	  

	     
	  
       
     def task_post(self,grab,task):
	  for elem in grab.doc.select('//div[@class="title"]/a'):
	       ur = grab.make_url_absolute(elem.attr('href'))
	       #print ur
	       if 'tvoyadres' in ur:
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)


     def task_item(self, grab, task):
	  try:
	       ray = grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]/a[contains(@href, "regiony")]').text()
	  except IndexError:
	       ray = ''          
	  try:
	       punkt= grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]/a[contains(@href, "goroda")]').text()
	  except IndexError:
	       punkt = ''
	  try:
	       ter= grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]').text()
	  except IndexError:
	       ter =''
	  try:
	       uliza = grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]/a[contains(@href, "ulitsy")]').text()
	  except IndexError:
	       uliza = ''
	  try:
	       dom = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Номер дома")]/following-sibling::dd').text()
	  except IndexError:
	       dom = ''
		
	  try:
	       orentir = grab.doc.select(u'//label[contains(text(),"Жилой комплекс:")]/following-sibling::p').text()
	  except IndexError:
	       orentir = ''              
	    
	  try:
	       metro = grab.doc.select(u'//span[contains(text(),"Кадастровый номер")]/following-sibling::span').text()
	    #print rayon
	  except IndexError:
	       metro = ''
	      
	  try:
	       metro_min = grab.doc.select(u'//span[contains(text(),"Отопление")]/following-sibling::span').text()
	  except IndexError:
	       metro_min = ''
	      
	  try:
	       metro_tr = grab.doc.select(u'//span[@class="object_item_metro_comment"]').text().split(u'мин. ')[1]
	  except IndexError:
	       metro_tr = ''

	  try:
	       try:
	            price = re.sub('[^\d\.]', u'',grab.doc.select(u'//title').text().split(u'цен')[1].split(u',')[0]).replace('.',' руб.')
	       except IndexError:
		    price = grab.doc.select(u'//meta[@name="description"]').attr('content').split(u'цене ')[1].split('.')[0]
	  except IndexError:
	       price = ''


	  try:
	       plosh_ob = grab.doc.select(u'//span[contains(text(),"Площадь земельного участка")]/following-sibling::span').text()
	  except IndexError:
	       plosh_ob = ''

	  
	       
	  try:
	       et = grab.doc.select(u'//th[contains(text(),"Газоснабжение")]/following-sibling::td').text()
	    #print price + u' руб'	    
	  except IndexError:
	       et = '' 
	      
	  try:
	       etagn = grab.doc.select(u'//dt[contains(text(),"Обновленно")]/following-sibling::dd[1]').text()
	    #print price + u' руб'	    
	  except IndexError:
	       etagn = ''

		
	  try:
	       opis = grab.doc.select(u'//meta[@name="description"]').attr('content') 
	  except IndexError:
	       opis = ''
	   
	  try:
               phone = re.sub('[^\d]', '',base64.b64decode(grab.doc.rex_text(u'data-tel=(.*?)==')[1:]+'=='))
          except IndexError:
	       phone = random.choice(list(open('../phone.txt').read().splitlines()))
	      
	  try:
	       try:
		    lico = grab.doc.rex_text(u'Собственник (.*?)"')
	       except IndexError: 
		    lico = grab.doc.select(u'//a[contains(@href, "polzovateli")]').text()
	  except IndexError:
	       lico = ''
	       
	  try:
	       comp = grab.doc.select(u'//span[contains(text(),"Сделка")]/following-sibling::span').text().replace(u'Продам',u'Продажа').replace(u'Сдам',u'Аренда')
	    #print rayon
	  except IndexError:
	       comp = ''
	       
	  try:
	       try:
		    try:
	                 data = grab.doc.select(u'//a[contains(@href, "uploads")]').attr('href').split('/uploads/')[1][:10].replace('/','.')
	            except IndexError:
	                 data = grab.doc.select(u'//div[@class="image"]/ul/li/img[contains(@src, "jpg")]').attr('src').split('/uploads/')[1][:10].replace('/','.') 
	       except IndexError:
		    data = grab.doc.select(u'//span[contains(text(),"Дата публикации")]/following-sibling::span').text()
	  except IndexError:
	       data = ''
	       
	       
	  try:
	       lat = grab.doc.select(u'//span[@id="map"]').attr('data-coordinates').split(',')[0]
	  except IndexError:
	       lat =''
	       
	  try:
	       lng = grab.doc.select(u'//span[@id="map"]').attr('data-coordinates').split(',')[1]
	  except IndexError:
	       lng =''		    

          data = reduce(lambda data, r: data.replace(r[0], r[1]), self.conv, data)
	  
	  projects = {'rayon': ray,
                      'punkt': punkt,
                      'teritor': ter,
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
                      'dol': lat,
                      'shir': lng,	                   
                      'lico':lico,
                      'company':comp,
                      'data':data}
	
	
	
	  yield Task('write',project=projects,grab=grab)

     def task_write(self,grab,task):
	  
	  print('*'*50)	       
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
	  
    
	  self.ws.write(self.result, 0,task.project['rayon'])
	  #self.ws.write(self.result, 1,task.project['rayon'])
	  self.ws.write(self.result, 2,task.project['punkt'])
	  self.ws.write(self.result, 33,task.project['teritor'])
	  self.ws.write(self.result, 4,task.project['ulica'])
	  self.ws.write(self.result, 5,task.project['dom'])
	  self.ws.write(self.result, 21,task.project['metro'])
	  self.ws.write(self.result, 19,task.project['naz'])
	  #self.ws.write(self.result, 10,task.project['object'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  #self.ws.write(self.result, 13, task.project['cena_m'])
	  #self.ws.write(self.result, 14, task.project['col_komnat'])
	  self.ws.write(self.result, 12, task.project['plosh_ob'])
	  #self.ws.write(self.result, 16, task.project['plosh_gil'])
	  self.ws.write(self.result, 31, task.project['dol'])
	  self.ws.write(self.result, 32, task.project['shir'])
	  self.ws.write(self.result, 15, task.project['etach'])
	  self.ws.write(self.result, 29, task.project['etashost'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'ТвойАдрес.RU')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 9, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	 
	  
	  print('*'*50)
	  print 'Ready - '+str(self.result)
	  print 'Tasks - %s' % self.task_queue.size()
	  print('*'*50)
	  self.result+= 1


bot = Nedvizhka_Zem(thread_number=5,network_try_limit=2000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=500, connect_timeout=500)
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
os.system("/home/oleg/pars/tvoy/comm.py")




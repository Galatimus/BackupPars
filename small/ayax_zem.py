#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
#from grab import Grab
import re
import os
import time
import xlsxwriter
from datetime import datetime,timedelta

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

workbook = xlsxwriter.Workbook(u'zemm/0001-0013_00_У_001-0081_AYAX.xlsx')  
    

class Ayax_Zem(Spider):
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
	  self.ws.write(0, 30, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(1,35):
               yield Task ('post',url='https://www.ayax.ru/zemelnye-uchastki/?PAGEN_1=%d'% x,network_try_count=100)
   
     def task_post(self,grab,task):
    
	  for elem in grab.doc.select(u'//ul[@class="photos-slideshow-proper"]/li[1]/a'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	 
        
        
     def task_item(self, grab, task):
	  try:
	       sub = u'Краснодарский край'#grab.doc.select(u'//label[contains(text(),"Адрес:")]/following-sibling::p').text().split(', ')[0]
	  except IndexError:
	       sub = ''
	  try:
	       ray = grab.doc.select(u'//div[contains(text(),"Район")]/preceding-sibling::div').text()
	     #print ray 
	  except IndexError:
	       ray = ''          
	  try:
	       punkt= grab.doc.select(u'//div[contains(text(),"Населенный пункт")]/preceding-sibling::div').text()
	  except IndexError:
	       punkt = ''
	       
	  try:
	       ter= ''#grab.doc.select(u'//p[@class="text-justify"]').text().split(', ')[0]
	  except IndexError:
	       ter =''
	       
	  try:
	       
	       uliza = grab.doc.select(u'//div[contains(text(),"Адрес")]/preceding-sibling::div').text().split(', ')[0]
	       #else:
		    #uliza = ''
	  except IndexError:
	       uliza = ''
	       
          try:
               dom = grab.doc.select(u'//div[contains(text(),"Адрес")]/preceding-sibling::div').text().split(', ')[1]
          except IndexError:
               dom = ''
	       
	  try:
	       trassa = grab.doc.select(u'//label[contains(text(),"Шоссе:")]/following-sibling::p').text().split(', ')[0]
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//span[@class="object__params-item_value"]').text()
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//div[@class="object__price"]').text()
	  except IndexError:
	       price = ''
	       
	  
	       
	  try:
	       plosh = grab.doc.select(u'//div[contains(text(),"Площадь")]/preceding-sibling::div').text()+u' м2'
	  except IndexError:
	       plosh = ''

	       
	  try:
	       vid = grab.doc.select(u'//div[contains(text(),"Тип участка")]/preceding-sibling::div').text()
	  except IndexError:
	       vid = '' 
	       
	       
	  try:
	       ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
	  except DataNotFound:
	       ohrana =''
	  try:
	       gaz = grab.doc.select(u'//div[contains(text(),"Газоснабжение")]/preceding-sibling::div').text()
	  except DataNotFound:
	       gaz =''
	  try:
	       voda =  grab.doc.select(u'//div[contains(text(),"Водоснабжение")]/preceding-sibling::div').text()
	  except DataNotFound:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//div[contains(text(),"Канализация")]/preceding-sibling::div').text()
	  except DataNotFound:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//div[contains(text(),"Электроснабжение")]/preceding-sibling::div').text()
	  except DataNotFound:
	       elek =''
	  try:
	       teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
	  except DataNotFound:
	       teplo =''
	       
	  try:
	       oper = u'Продажа'#grab.doc.select(u'//li[@class="c-header__menu__item  c-header__menu__item_active"]').text() 
	  except DataNotFound:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//div[@itemprop="description"]').text()#.replace(u'Описание ','') 
	  except IndexError:
	       opis = ''
	       
	  try:
	       phone = re.sub('[^\d\+]','',grab.doc.rex_text(u'tel:(.*?)"'))
	  except IndexError:
	       phone = ''
	       
	  try:
	       lico = grab.doc.select(u'//div[@class="employee__info_link"]').text()
	  except IndexError:
	       lico = ''
	       
	  try:
	       comp = u'АЯКС-Риэлт'# grab.doc.select(u'//img[@class="thumbnail"][contains(@src,"agency")]/following::b[1]').text()
	  except IndexError:
	       comp = ''
	       
	  try:
	       
	       conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
	                 (u' мая ',u'.05.'),(u' июня ',u'.06.'),
	                 (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
	                 (u' января ',u'.01.'),(u' декабря ',u'.12.'),
	                 (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
	                 (u' февраля ',u'.02.'),(u' октября ',u'.10.')] 
	       dt= grab.doc.select(u'//b[contains(text(),"Обновлено:")]/following-sibling::span').text()
	       data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(' ','')#.replace(u'более3-хмесяце', u'07.2015')
	    #print data
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
                      'vid': vid,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'phone':phone,
                      'lico':lico,
                      'company':comp,
                      'data':data,
                      'oper':oper
                      }
          
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
	  self.ws.write(self.result, 3, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  #self.ws.write(self.result, 3, task.project['teritor'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 7, task.project['trassa'])
	  self.ws.write(self.result, 11, task.project['udal'])
	  self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 14, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 19, task.project['teplo'])
	  self.ws.write(self.result, 20, task.project['ohrana'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'ГК "Аякс риэлт"')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	  #self.ws.write(self.result, 30, task.project['sub']+
	                #', '+task.project['punkt']+
	                #', '+task.project['rayon']+
	                #', '+task.project['ulica']+
	                #', '+task.project['dom'])
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  #print oper
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  #if self.result > 20:
	       #self.stop()

     
bot = Ayax_Zem(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
workbook.close()
print('Done!') 

time.sleep(5)
#os.system("/home/oleg/pars/small/city_zem.py")







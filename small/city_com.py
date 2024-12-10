#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
from sub import conv

#from PIL import Image
import os
import math
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






   
     
class move_Com(Spider):
     def prepare(self):
	  self.workbook = xlsxwriter.Workbook(u'comm/0001-0002_00_C_001-0033_CITYST.xlsx')
	  self.ws = self.workbook.add_worksheet()
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
	  self.ws.write(0, 17, u"ГОД_ПОСТРОЙКИ")
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
	  self.ws.write(0, 32, u"КАДАСТРОВЫЙ_НОМЕР")
	  self.ws.write(0, 33, u"ЗАГОЛОВОК")
	  self.ws.write(0, 34, u"ШИРОТА_ИСХ")
	  self.ws.write(0, 35, u"ДОЛГОТА_ИСХ")
	  self.conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
                  (u' мая ',u'.05.'),(u' июня ',u'.06.'),
                  (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
                  (u' января ',u'.01.'),(u' декабря ',u'.12.'),
                  (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
                  (u' февраля ',u'.02.'),(u' октября ',u'.10.'),
                  (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
                  (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]	       
	 
	  self.result= 1
	  
	   
	   
	   
	   
     def task_generator(self):
	  yield Task ('next',url='http://magnitogorsk-citystar.ru/change-city',refresh_cache=True,network_try_count=100)
	  
	  
     def task_next(self,grab,task):
	  for el in grab.doc.select(u'//h1/following-sibling::ul/li/a'):
	       urr = grab.make_url_absolute(el.attr('href'))  
	       #print urr
	       yield Task('post', url=urr+'/realty/prodazha-komm-nedvizhimosti/',refresh_cache=True,network_try_count=100)
	       yield Task('post', url=urr+'/realty/sdacha-komm-nedvizhimosti/',refresh_cache=True,network_try_count=100)
     
   
     def task_post(self,grab,task):
	  for elem in grab.doc.select('//a[@class="detail-link"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
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
	       mesto = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text()
	  except IndexError:
	       mesto =''
	       
	  try:
	       punkt = grab.doc.select(u'//div[@class="cur-city-name"]').text().title()
	  except IndexError:
	       punkt = ''	       
	   
	  try:
	       ter= grab.doc.select(u'//td[contains(text(),"Район")]/following-sibling::td').text()
	       
	  except IndexError:
	       ter =''
	  try:
	      
	       uliza= grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text()		    
	  except (IndexError,UnboundLocalError):
	       uliza =''
	  try:
	       dom = grab.doc.select(u'//h1').text()		  
	  except (IndexError,AttributeError):
	       dom = ''
	    
	  try:
	       tip = grab.doc.select(u'//td[contains(text(),"Вид недвижимости")]/following-sibling::td').text()
	  except IndexError:
	       tip = ''
	  try:
	       naz = grab.doc.select(u'//span[contains(text(),"Назначение")]/following::div[2]').text()
	  except IndexError:
	       naz =''
	  try:
	       klass =  grab.doc.select(u'//span[contains(text(),"Этаж")]/following::div[2]').text()
	  except IndexError:
	       klass = ''
	  try:
	       
	       price = grab.doc.select(u'//td[contains(text(),"Цена")]/following-sibling::td/span').text()+u' р.'		    
	  except IndexError:
	       price =''
	  try: 
	       plosh = grab.doc.select(u'//td[contains(text(),"Площадь помещения")]/following-sibling::td').text()
	  except IndexError:
	       plosh=''
	  try:
	       ohrana = grab.doc.select(u'//span[contains(text(),"Этажность")]/following::div[2]').text()
	  except IndexError:
	       ohrana =''
	  try:
	       gaz =  grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Материал стен")]/following-sibling::dd').text()
	  except IndexError:
	       gaz =''
	  try:
	       voda =  grab.doc.select(u'//h4').text().split('/')[1]
	  except IndexError:
	       voda =''
	  try:
	       d1 = grab.doc.select(u'//div[@class="date"][2]').text().replace(u'Дата подачи: ','').split(u'г.')[0]
	       kanal = reduce(lambda d1, r: d1.replace(r[0], r[1]), self.conv, d1)
	  except IndexError:
	       kanal =''
	  try:
	       elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
	  except DataNotFound:
	       elek =''
	  try:
	       teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
	  except DataNotFound:
	       teplo =''
	  #time.sleep(1)
	  try:
	       opis = grab.doc.select(u'//td[@class="note"]').text() 
	  except IndexError:
	       opis = ''
	  try:
	       lico = grab.doc.select(u'//div[@class="phone"]/preceding-sibling::div[@class="name"]').text()
	  except IndexError:
	       lico = ''
	  try:
	       comp = grab.doc.select(u'//span[contains(text(),"Агентство недвижимости:")]/following-sibling::div[1]').text()
	  except IndexError:
	       comp = ''
	  
	  try: 
	       
	       d = grab.doc.select(u'//div[@class="date"][1]').text().replace(u'Дата подачи: ','').split(u'г.')[0]
	       data = reduce(lambda d, r: d.replace(r[0], r[1]), self.conv, d)		   
	  except IndexError:
	       data=''
	       
	  
	  try:
	       phone = re.sub('[^\d\,]','',grab.doc.select(u'//span[contains(text(),"тел.:")]/following-sibling::text()').text())
	  except IndexError:
	       phone = ''
	       
	  try:
	       if 'prodazha' in task.url:
	            oper = u'Продажа' 
	       elif 'sdacha' in task.url:
	            oper = u'Аренда'
	       else:
	            oper = ''
          except IndexError:
	       oper = ''           
	  
	  sub = reduce(lambda punkt, r: punkt.replace(r[0], r[1]), conv, punkt)

   
	  projects = {'sub': sub,
                     'adress': mesto,
                      'terit':ter, 
                      'punkt':punkt, 
                      'ulica':uliza,
                      'dom':dom,
                      'tip':tip,
                      'naz':naz,
                      'klass': klass,
                      'cena': price,
                      'plosh': plosh,
	              'operacia': oper,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': re.sub('[^\d\.]','',kanal),
                      'electr': elek,
                      'teplo': teplo,
                      'opis': opis,
                      'url': task.url,
                      'phone': phone,
                      'lico':lico,
                      'company': comp,
                      'data':re.sub('[^\d\.]','',data)}
		      
     
     
	  yield Task('write',project=projects,grab=grab)
     
     
     
     def task_write(self,grab,task):
	  #time.sleep(1)
	  print('*'*100)	       
	  print  task.project['sub']
	  print  task.project['punkt']
	  print  task.project['adress']
	  print  task.project['terit']
	  print  task.project['ulica']
	  print  task.project['dom']
	  print  task.project['tip']
	  print  task.project['naz']
	  print  task.project['klass']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanaliz']
	  print  task.project['electr']
	  print  task.project['teplo']
	  print  task.project['ohrana']
	  print  task.project['opis']
	  print  task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	 
	 
     
	  
	  

	  self.ws.write(self.result, 0, task.project['sub'])
	  #self.ws.write(self.result, 1, task.project['adress'])
	  self.ws.write(self.result, 3, task.project['terit'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 33, task.project['dom'])
	  self.ws.write(self.result, 9, task.project['tip'])
	  #self.ws.write(self.result, 9, task.project['naz'])
	  #self.ws.write(self.result, 13, task.project['klass'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  #self.ws.write(self.result, 14, task.project['ohrana'])
	  #self.ws.write(self.result, 16, task.project['gaz'])
	  #self.ws.write(self.result, 24, task.project['voda'])
	  self.ws.write(self.result, 30, task.project['kanaliz'])
	  #self.ws.write(self.result, 23, task.project['electr'])
	  #self.ws.write(self.result, 24, task.project['teplo'])
	  self.ws.write(self.result, 18, task.project['opis'])
	  self.ws.write(self.result, 19, u'СИТИСТАР')
	  self.ws.write_string(self.result, 20, task.project['url'])
	  self.ws.write(self.result, 21, task.project['phone'])
	  self.ws.write(self.result, 22, task.project['lico'])
	  self.ws.write(self.result, 23, task.project['company'])
	  self.ws.write(self.result, 29, task.project['data'])
	  #self.ws.write(self.result, 32, task.project['data1'])
	  self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 28, task.project['operacia'])
	  self.ws.write(self.result, 24, task.project['sub']+
                        ', '+task.project['punkt']+
                        ', '+task.project['terit']+
                        ', '+task.project['adress'])	       	       
	  print('*'*10)
	  #print self.sub
	  print 'Ready - '+str(self.result)
	  logger.debug('Tasks - %s' % self.task_queue.size()) 
	  print task.project['operacia']
	  print('*'*10)
	  self.result+= 1
	  
	 
	  
	  
	  
	  #if self.result >= 50:
	       #self.stop()	       

bot = move_Com(thread_number=5, network_try_limit=2000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
bot.run()
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
bot.workbook.close()
    
       
     
     
     
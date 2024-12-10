#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import os
import math
import requests
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






workbook = xlsxwriter.Workbook(u'0001-0002_00_У_001-0027_NERS.xlsx')





class Ners_zem(Spider):
     def prepare(self):
	  self.ws = workbook.add_worksheet(u'Ners__Земля')
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
	  self.conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
               (u' мая ',u'.05.'),(u' июня ',u'.06.'),
               (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
               (u' января ',u'.01.'),(u' декабря ',u'.12.'),
               (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
               (u' февраля ',u'.02.'),(u' октября ',u'.10.'),
               (u'сегодня,', (datetime.today().strftime('%d.%m.%Y'))),
               (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]	       
	 
	  self.result= 1
	  
	   
	   
	   
	   
     def task_generator(self):
	  l= open('ners_zem1.txt').read().splitlines()
	  self.dc = len(l)
	  print self.dc
	  for line in l:
	       yield Task ('item',url=line,refresh_cache=True,network_try_count=100)

     def task_item(self, grab, task):
	  try:
	       sub = grab.doc.select(u'//ul[@class="linklist navlinks"]/li[1]/a').text()
	  except IndexError:
	       sub = ''	  	  
	  try:
	       mesto = grab.doc.select(u'//dt[contains(text(),"Район")]/following-sibling::dd').text()
	  except IndexError:
	       mesto =''
	       
	  try:
	       if sub == u"Москва":
		    punkt= u"Москва"
	       elif sub == u"Санкт-Петербург":
		    punkt= u"Санкт-Петербург"
	       elif sub == u"Севастополь":
		    punkt= u"Севастополь"
	       else:
		    punkt = grab.doc.select(u'//dt[contains(text(),"Населенный пункт:")]/following-sibling::dd').text()
	  except IndexError:
	       punkt = ''	       
	   
	  try:
	       ter =  grab.doc.select(u'//dt[contains(text(),"Шоссе:")]/following-sibling::dd').text()
	  except IndexError:
	       ter =''
	  try:
	       ul = grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text()
	       if ' ул'in ul:
		    uliza = ul.split(', ')[0]
		    tip = ''
	       else:
		    tip = ul.split(', ')[0]
		    uliza = ''
	  except IndexError:
	       uliza = ''
	       tip = ''
	  try:
	       dom = re.sub('[^\d]','',grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text().split(', ')[1])
	  except IndexError:
	       dom = ''
	       
	  
	  try:
	       naz = grab.doc.select(u'//span[contains(text(),"Использование")]/following::div[2]').text()
	  except IndexError:
	       naz =''
	  try:
	       klass =  grab.doc.select(u'//span[contains(text(),"Расстояние до города")]/following::div[2]').text()
	  except IndexError:
	       klass = ''
	  try:
	       price = grab.doc.select(u'//dt[contains(text(),"Цена:")]/following-sibling::dd/text()').text()
	  except IndexError:
	       price =''
	  try: 
	       plosh = grab.doc.select(u'//dt[contains(text(),"Площадь:")]/following-sibling::dd').text()
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
	       voda =  grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text()
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//dt[contains(text(),"Цена:")]/following-sibling::dd/span').text().replace(u' за сотку)','').replace('(','')
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
	       opis = grab.doc.select(u'//div[@class="param"]/following-sibling::div[@class="info"]').text() 
	  except IndexError:
	       opis = ''
	  try:
	       lico = grab.doc.select(u'//a[@class="profile_link"]').text()
	  except IndexError:
	       lico = ''
	  try:
	       comp = grab.doc.select(u'//a[@class="firm_link"]').text()
	  except IndexError:
	       comp = ''
	  try:
	       d = grab.doc.select(u'//div[contains(text(),"Дата размещения:")]').text().replace(u'Дата размещения: ','') 
	       data1 = reduce(lambda d, r: d.replace(r[0], r[1]), self.conv, d)
	  except IndexError:   
	       data1 = ''
	  try: 
	       dt = grab.doc.select(u'//div[contains(text(),"Дата обновления:")]').text().replace(u'Дата обновления: ','')#[:9]
	       data = reduce(lambda dt, r: dt.replace(r[0], r[1]), self.conv, dt)[:10]
	  except IndexError:
	       data=''
	       

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
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis': opis,
                      'url': task.url,
                      'lico':lico,
                      'company': comp,
                      'data':data,
                      'data1':data1}
	 
	    
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
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	  print  task.project['data1']
	 
     
	  
	  

	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['adress'])
	  self.ws.write(self.result, 7, task.project['terit'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 3, task.project['tip'])
	  self.ws.write(self.result, 14, task.project['naz'])
	  self.ws.write(self.result, 8, task.project['klass'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 20, task.project['ohrana'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 31, task.project['voda'])
	  self.ws.write(self.result, 11, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  #self.ws.write(self.result, 24, task.project['teplo'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'Национальная единая риэлторская сеть')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  #self.ws.write(self.result, 25, task.phone)
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data1'])
	  self.ws.write(self.result, 29, task.project['data'])
	  self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 9, u'Продажа')
	  print('*'*100)
	  
	  print 'Ready - '+str(self.result)+'/'+str(self.dc)
	  logger.debug('Tasks - %s' % self.task_queue.size()) 
	  print('*'*100)
	  self.result+= 1
	  
	 
	  
	  
	  
	  #if self.result > 10:
	       #self.stop()	       
   

    

bot = Ners_zem(thread_number=3, network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
#bot.create_grab_instance(timeout=500, connect_timeout=5000)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
command = 'mount -a'
os.system('echo %s|sudo -S %s' % ('1122', command))
time.sleep(3)
workbook.close()
print('Done')
   
       
     
     
     
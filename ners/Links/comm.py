#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import math
import re
import time
import random
import os
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






workbook = xlsxwriter.Workbook(u'0001-0002_00_C_001-0027_NERS.xlsx')






class QP_Com(Spider):
     def prepare(self):
	  
	  self.ws = workbook.add_worksheet(u'ners')
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
               (u'сегодня,', (datetime.today().strftime('%d.%m.%Y'))),
               (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]
	  self.result= 1
	  
	   
	   
	   
	   
     def task_generator(self):
	  l= open('ners_com.txt').read().splitlines()
	  self.dc = len(l)
	  print self.dc
	  for line in l:
	       yield Task ('item',url=line,refresh_cache=True,network_try_count=50)


     def task_item(self, grab, task):
	  try:
	       sub = grab.doc.select(u'//div[@id="siteMenu"]/ul/li[1]/a').text()
	  except IndexError:
	       sub = ''	  
	  try:
	       mesto = grab.doc.select(u'//dt[contains(text(),"Район области:")]/following-sibling::dd').text()
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
	       ter =  grab.doc.select(u'//dt[contains(text(),"Шоссе:")]/following-sibling::dd').text()+' ш.'
	  except IndexError:
	       ter =''
	  try:
	       uliza = grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text().split(', ')[0]
	  except IndexError:
	       uliza = ''
	  try:
	       dom = re.sub('[^\d]','',grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text().split(', ')[1])
	  except IndexError:
	       dom = ''
	       
	  try:
	       tip = grab.doc.select(u'//dt[contains(text(),"Метро:")]/following-sibling::dd').text()
	  except IndexError:
	       tip = ''
	  try:
	       naz = grab.doc.select(u'//dt[contains(text(),"Тип объекта:")]/following-sibling::dd').text()
	  except IndexError:
	       naz =''
	  try:
	       klass =  grab.doc.select(u'//dt[contains(text(),"Этаж:")]/following-sibling::dd').text()
	  except IndexError:
	       klass = ''
	  try:
	       price = grab.doc.select(u'//div[@class="price_value"]/text()').text()
	  except IndexError:
	       price =''
	  try: 
	       plosh = grab.doc.select(u'//dt[contains(text(),"Общая площадь:")]/following-sibling::dd').text()
	  except IndexError:
	       plosh=''
	  try:
	       ohrana = grab.doc.select(u'//dt[contains(text(),"Этажность:")]/following-sibling::dd').text()
	  except IndexError:
	       ohrana =''
	  try:
	       gaz =  grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text()
	  except IndexError:
	       gaz =''
	  try:
	       voda =  grab.doc.select(u'//dt[contains(text(),"Метро:")]/following-sibling::dd').text()
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//h1/span').text().split(' ')[0]
	  except IndexError:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//dt[contains(text(),"До метро:")]/following-sibling::dd').text()
	  except IndexError:
	       elek =''
	  try:
	       teplo = grab.doc.select(u'//title').text()
	  except IndexError:
	       teplo =''
	  #time.sleep(1)
	  try:
	       opis = grab.doc.select(u'//div[@class="info mb-4"]').text() 
	  except IndexError:
	       opis = ''
	       
	  #for p in range(1,51):
	       #try:
		    #key = grab.doc.select(u'//div[@class="notes_id"]/b').text() 
		    #user = grab.doc.select(u'//div[@id="get_phone"]').attr('data-u')
		    #pkey = grab.doc.select(u'//body/@data-stat-id').text()
		    #db = re.sub('[^\d]','',grab.doc.select(u'//article[@id="notes_wrap"]/@data-db_importer_id').text())
		    #my = re.sub('[^\d]','',grab.doc.select(u'//article[@id="notes_wrap"]/@data-is_my_notes').text())
		    #link = task.url.split(u'object')[0]#+u'ru'
		    #url_ph = 'https://ru.ners.ru/ajax/?module=notes_get_phone&notes_id='+key+'&user_id='+user+'&db_importer_id='+db+'&stat_id='+pkey+'&is_my_notes='+my+'&app_alias=ru&json=1'
		    #headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
			      #'Accept-Encoding': 'gzip, deflate, br',
			      #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			      #'Host': 'ru.ners.ru',
			      #'Origin': link,		              
			      #'Referer': task.url,
			      #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0'}
		    #g2 = grab.clone(headers=headers,proxy_auto_change=True)
		    #time.sleep(1)
		    ##g2.request(post=[('app_alias','ru'), ('db_importer_id', db),('is_my_notes', my),('json','1'), ('module', 'notes_set_view'),('notes_id', key),('stat_id', pkey)],headers=headers,url=url_ph)
		    #g2.go(url_ph)
		    #g2.set_input_by_xpath('//div[@id="get_phone"]', 'bar-value')
		    #g2.submit(make_request=False)
		    #print g2.response.body
		    #print 'Phone-OK'
		    #del g2
		    #break 		    
	       #except (IndexError,GrabConnectionError,GrabNetworkError,GrabTimeoutError):
		    #print 'Change proxy'+' : '+str(p)+' / 31'
		    #g2 = grab.clone(headers=headers,timeout=2, connect_timeout=2,proxy_auto_change=True)

	       
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
	       
	  try:
               lat = grab.doc.rex_text(u'notes_coord = (.*?)]').split(', ')[0][1:]
          except IndexError:
	       lat = ''
	  
	  try:
	       lng = grab.doc.rex_text(u'notes_coord = (.*?)]').split(', ')[1]
	  except IndexError:
	       lng = ''	       
	  

   
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
                      'oper': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis': opis,
	              'phone': random.choice(list(open('../phone.txt'))),
                      'url': task.url,
                      'lico':lico,
                      'company': comp,
	              'shir':lat,
	              'dol':lng,	              
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
	  #print  task.project['voda']
	  print  task.project['electr']
	  print  task.project['teplo']
	  print  task.project['ohrana']
	  print  task.project['opis']
	  print  task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	  print  task.project['data1']
	  print  task.project['shir']
	  print  task.project['dol']	  
	 
     
	  
	  

	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['adress'])
	  self.ws.write(self.result, 6, task.project['terit'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 26, task.project['tip'])
	  self.ws.write(self.result, 9, task.project['naz'])
	  self.ws.write(self.result, 15, task.project['klass'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  self.ws.write(self.result, 16, task.project['ohrana'])
	  self.ws.write(self.result, 24, task.project['gaz'])
	  self.ws.write(self.result, 21, task.project['phone'])
	  self.ws.write(self.result, 34, task.project['shir'])
	  self.ws.write(self.result, 27, task.project['electr'])
	  self.ws.write(self.result, 33, task.project['teplo'])
	  self.ws.write(self.result, 18, task.project['opis'])
	  self.ws.write(self.result, 19, u'Национальная единая риэлторская сеть')
	  self.ws.write_string(self.result, 20, task.project['url'])
	  self.ws.write(self.result, 35, task.project['dol'])
	  self.ws.write(self.result, 22, task.project['lico'])
	  self.ws.write(self.result, 23, task.project['company'])
	  self.ws.write(self.result, 30, task.project['data'])
	  self.ws.write(self.result, 29, task.project['data1'])
	  self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 28, task.project['oper'])
	  print('*'*100)
	  #print self.sub
	  print 'Ready - '+str(self.result)+'/'+str(self.dc)
	  logger.debug('Tasks - %s' % self.task_queue.size()) 
	  print  task.project['oper']
	  print('*'*100)
	  self.result+= 1
	  
	 
	  #if int(self.result) == int(self.num):
	       #self.stop()	       
	  
	  
	  #if self.result > 100:
	       #self.stop()	       


bot = QP_Com(thread_number=5, network_try_limit=500)
bot.load_proxylist('../ivan.txt','text_file')
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
  
     
     
     
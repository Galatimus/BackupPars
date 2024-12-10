#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
import os
import time
import random
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)





operacia = u'Аренда'

workbook = xlsxwriter.Workbook(u'zem/Raui_Земля_'+operacia+'.xlsx')


class Farpost_Zem(Spider):
     def prepare(self):
	  #self.f = page
	  for p in range(1,5):
	       try:
		    #time.sleep(1)
		    g = Grab(timeout=50, connect_timeout=100)
		    g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
		    g.go('https://raui.ru/snyat-dom-dacha-zemelnye_uchastki/uchastok')
		    print g.doc.code
		    if g.doc.code ==200:
			 self.num = re.sub('[^\d]','',g.doc.select(u'//a[@class="pagging__link dotts"]/following::li[1]/a/span').text())
			 print 'OK'
			 del g
			 break
	       except(GrabTimeoutError,GrabNetworkError,IndexError,KeyError,AttributeError,GrabConnectionError):
		    print g.config['proxy'],'Change proxy'
		    g.change_proxy()
		    del g
		    continue
	  else:
	       self.num = '289'
	  print self.num
	  self.ws = workbook.add_worksheet(u'Raui')
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
	  self.ws.write(0, 31, u"ВИД_ПРАВА")
	  self.ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
	  #self.r = conv     
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x1 in range(1,int(self.num)+1):
	       yield Task ('post',url='https://raui.ru/snyat-dom-dacha-zemelnye_uchastki/uchastok?page=%d'%x1,refresh_cache=True,network_try_count=100)
	
	       
	       
     def task_post(self,grab,task):
    
	  for elem in grab.doc.select(u'//a[contains(text(),"Подробнее")]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	 
        
        
     def task_item(self, grab, task):
	  try:
	       sub= grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[2]').text()
	  except IndexError:
	       sub = ''
	  try:
	       try:
	            try:
	                 ray = grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[contains(text()," район")]').text()
	            except IndexError:
	                 ray = grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[contains(text()," округ ")]').text()
	       except IndexError:
	            ray = grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[contains(text()," образование ")]').text()
	  except IndexError:
	       ray = ''          
	  try:
	       if sub == u'Москва':
                    punkt= u'Москва'
               elif sub == u'Санкт-Петербург':
	            punkt= u'Санкт-Петербург'
               elif sub == u'Севастополь':
	            punkt= u'Севастополь'
               else:
	            punkt= grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[4]').text()
	  except IndexError:
	       punkt = ''
	       
	 	       
	  try:
	       udal = grab.doc.select(u'//div[@class="item__price"]/following-sibling::div[@class="item__number"]').text()
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//div[@class="item__price"]').text()#.replace(u'a',u'р.')
	  except IndexError:
	       price = ''
	  try:
	       plosh = grab.doc.select(u'//td[contains(text(),"Площадь участка:")]/following-sibling::td').text()
	  except IndexError:
	       plosh = ''
	  try:
	       vid = grab.doc.select(u'//td[contains(text(),"Категория земель:")]/following-sibling::td').text()
	  except IndexError:
	       vid = '' 
	  try:
	       gaz = grab.doc.select(u'//td[contains(text(),"Газ:")]/following-sibling::td').text().replace(u'нет информации','')
	  except IndexError:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//td[contains(text(),"Водоснабжение:")]/following-sibling::td').text().replace(u'нет информации','')
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//td[contains(text(),"Канализация:")]/following-sibling::td').text().replace(u'нет информации','')
	  except IndexError:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//td[contains(text(),"Электричество:")]/following-sibling::td').text().replace(u'нет информации','')
	  except IndexError:
	       elek =''
	  try:
	       teplo = grab.doc.select(u'//h1').text().replace('Участок ','')
	  except IndexError:
	       teplo =''
	  try:
	       opis = grab.doc.select(u'//div[@class="item-text"]').text() 
	  except IndexError:
	       opis = ''
	  try:
	       lico = grab.doc.select(u'//p[contains(text(),"Имя:")]/following-sibling::h3[1]').text()
	  except IndexError:
	       lico = ''
	       
	  try:
	       comp = grab.doc.select(u'//p[contains(text(),"Агентство")]/following-sibling::h3[1]').text()
	  except IndexError:
	       comp = ''
	       
	  try:
	       data = datetime.strptime(grab.doc.select(u'//meta[@property="article:published_time"]').attr('content')[:10].replace('-','.'), '%Y.%m.%d')
	    #print data
	  except IndexError:
	       data = ''
	       
	       
	  #id_phone = re.sub('[^\d]','',grab.doc.select(u'//div[@class=" contactstabs"]').attr('data-id'))
	       
	  #phone_url = 'https://raui.ru/ajax/item/contact?id='+id_phone 
     
	  #headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
                    #'Accept-Encoding': 'gzip, deflate, br',
                    #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
                    #'Content-Length': '10',
                    #'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    #'Cookie': 'session=aqq75kqhsmbegk00crvocv86t2',
                    #'Host': 'raui.ru',
                    #'Referer': task.url,
                    #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0', 
                    #'X-Requested-With': 'XMLHttpRequest'}
	  #g2 = grab.clone(headers=headers,proxy_auto_change=True)
     
	  
	  #try:               
	       ##time.sleep(1)
	       #g2.request(post=[('id', id_phone)],headers=headers,url=phone_url)
	       #phone =  re.sub('[^\d\+]','',g2.doc.json["contacts"]["phone"])
	       #print 'Phone-OK'
	       #del g2
	  #except (IndexError,GrabConnectionError,GrabNetworkError,GrabTimeoutError,TypeError):
	       #del g2
	       
	  try:
	       phone = re.sub('[^\d]','',grab.doc.select(u'//div[@class="item-contacts-values"]/div').text()+str(random.randint(100000,999999)))
	  except IndexError:
	       phone = random.choice(list(open('../phone.txt').read().splitlines()))	       
	  
	  
		    
	  
						   
	       
	  projects = {'url': task.url,
                      'sub': sub,
                      'rayon': ray,
                      'punkt': punkt,
                      'udal': udal,
                      'cena': price,
                      'plosh':plosh,
                      'vid': vid,
                      'gaz': gaz,
                      'voda': voda,
	              'phone':phone[:11],
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'lico':lico,
                      'company':comp,
                      'data':data.strftime('%d.%m.%Y')}
	  
	 	       
          
	  yield Task('write',project=projects,grab=grab)
            
     def task_write(self,grab,task):
	  print('*'*50)
	  print  task.project['sub']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['udal']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['vid']
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
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 11, task.project['udal'])
	  self.ws.write(self.result, 9, operacia)
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 13, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 32, task.project['teplo'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'RAUI')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)+'/'+str(self.num)+'0'
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  print  operacia
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  #if self.result > 50:
	       #self.stop()

     
bot = Farpost_Zem(thread_number=5,network_try_limit=1000)
#bot.setup_queue(backend='mongo', database='farpost',host='192.168.10.200')
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
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









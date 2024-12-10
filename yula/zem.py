#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import random
import os
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



i = 0
l= open('zem.txt').read().splitlines()
page = l[i] 
oper = u'Продажа'




while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Theproperty_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       for p in range(1,11):
		    try:
			 time.sleep(1)
			 g = Grab(timeout=10, connect_timeout=15)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http') 
			 g.go(self.f)
			 self.dt = g.doc.select(u'//span[@class="product_item__location"]').text()
			 print self.dt
			 link_sub = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+'Росcия, '+self.dt
			 time.sleep(1)
			 g.go(link_sub) 
			 self.sub = g.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["Address"]["Components"][2]["name"]
			 time.sleep(1)
			 print self.sub
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,IndexError,KeyError,AttributeError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
	            self.sub = ''
		    self.stop()


	       
	       self.workbook = xlsxwriter.Workbook(u'zem/youla_zem'+'_'+str(i+1)+'.xlsx')
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
	       self.ws.write(0, 30, u"ЗАГОЛОВОК")
	       self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):
	       yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
	          
	       
            
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//li[@class="product_item"]/a'):
                    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	       
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//div[@class="pagination__button"]/a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print('*'*100)
		    print '!!!','NO PAGE NEXT','!!!'
		    print('*'*100)	  
     
	  def task_item(self, grab, task):
	       try:
		    ray = grab.doc.rex_text(u'isFavorite(.*?)city').decode("unicode_escape").split(u'description')[1][3:].split(u'latitude')[0][:-3]
	       except IndexError:
		    ray = ''          
	       try:
	            punkt = self.dt
	       except IndexError:
		    punkt = ''

	       try:
		    metro = grab.doc.select(u'//title').text().split(', ')[2].split('купить ')[0][:-2]
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//h1').text().split(u' — ')[1]
		 #print rayon
	       except IndexError:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//span[@class="object_item_metro_comment"]').text().split(u'мин. ')[1]
	       except IndexError:
		    metro_tr = ''

	       try:
	            price = grab.doc.select(u'//title').text().split(u'цена ')[1].split(u'руб.')[0]+' руб.'
               except IndexError:
                    price = ''

	       try:
		    plosh_ob = grab.doc.select(u'//title').text().split(', ')[1]#.split('купить ')[0][:-2]
		  #print rayon
	       except IndexError:
		    plosh_ob = ''
     
	       
		    
	       try:
		    et = grab.doc.select(u'//th[contains(text(),"Газоснабжение")]/following-sibling::td').text()
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		    etagn = grab.doc.select(u'//dt[contains(text(),"Тип сделки")]/following-sibling::dd[1]').text()
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''
		     
	       
		
		     
	       try:
		    opis = grab.doc.rex_text(u'Moscow area"}},"offset"(.*?)","price').decode("unicode_escape").split(u'MSK+00 - Moscow area')[1].split(u'description')[1][3:]
	       except IndexError:
	            opis = ''
		
	       try:
                    phone = grab.doc.rex_text(u'displayPhoneNum":"(.*?)"')
               except IndexError:
	            phone = random.choice(list(open('../phone.txt').read().splitlines()))
		   
	       try:
		    lico = grab.doc.rex_text(u'Moscow area"}},"offset"(.*?)","price').decode("unicode_escape").split(u'isOnline')[0].split(u'name')[1][3:-3]
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//title').text()
	       except IndexError:
		    comp = ''
		    
	       try:
		    data = grab.doc.select(u'//title').text().split('дата размещения: ')[1].split(' Продажа')[0][:-2]
	       except IndexError:
		    data = ''

	   
	       projects = {'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
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
	       
	       try:
		    link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ray
		    yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
			 yield Task('adres',grab=grab,project=projects)	  

	  def task_adres(self, grab, task):
     
	       try:
		    ter=  grab.doc.rex_text(u'SubAdministrativeAreaName":"(.*?)"')
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
     
	       project2 ={'teritor': ter,
	                  'ulica':uliza,
	                  'dom':dom.replace('/','')} 
	     
	     
	       yield Task('write',project=task.project,proj=project2,grab=grab)
	     
	     
	     
	     
	     
	     
	  def task_write(self,grab,task):
	       
	       print('*'*50)	       
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.proj['teritor']
               print  task.proj['ulica']
               print  task.proj['dom']	 
	       print  task.project['metro']
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
	       print  task.project['naz']	      
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 31,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 1,task.proj['teritor'])
	       self.ws.write(self.result, 4,task.proj['ulica'])
	       self.ws.write(self.result, 5,task.proj['dom'])
	       self.ws.write(self.result, 14,task.project['metro'])
	       #self.ws.write(self.result, 32,task.project['naz'])
	       #self.ws.write(self.result, 10,task.project['object'])
	       self.ws.write(self.result, 9,oper)
	       self.ws.write(self.result, 10, task.project['cena'])
	       #self.ws.write(self.result, 13, task.project['cena_m'])
	       #self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 12, task.project['plosh_ob'])
	       #self.ws.write(self.result, 16, task.project['plosh_gil'])
	       #self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       #self.ws.write(self.result, 18, task.project['plosh_com'])
	       self.ws.write(self.result, 15, task.project['etach'])
	       self.ws.write(self.result, 30, task.project['etashost'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'YOULA.RU')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 30, task.project['company'])
	       self.ws.write(self.result, 28, task.project['data'])
	       self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	      
	       
	       print('*'*50)
	       print self.sub
	      
	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*50)
	       self.result+= 1
	       
	       
	       
	       
	       #if self.result > 10:
		    #self.stop()

	 
     
     bot = Theproperty_Zem(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')  
     command = 'mount -a'
     os.system('echo %s|sudo -S %s' % ('1122', command))
     time.sleep(2)
     bot.workbook.close()
     print('Done')
     del bot
     
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break

     
     

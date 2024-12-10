#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import time
import math

from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



i = 0
l= open('Links/Zem_prod.txt').read().splitlines()

page = l[i] 
oper = u'Продажа'





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Nndv_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       #self.link =l[i]
	       while True:
		    try:
			 time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='/home/oleg/Proxy/tipa.txt',proxy_type='http')
			 g.go(self.f)
			 
			 self.sub = g.doc.select(u'//div[@class="user-menu"]/ul[2]').text()
			 self.num = g.doc.select(u'//div[@data-test="offers-list-header-results"]').number()
                         self.pag = int(math.ceil(float(int(self.num))/float(25)))
                         
			 print self.sub,self.pag,self.num
			 del g
			 break
                    except(GrabTimeoutError,GrabNetworkError,IndexError, GrabConnectionError):
                         print g.config['proxy'],'Change proxy'
                         g.change_proxy()
			 del g
                         continue
                    
		    
	       
	       self.workbook = xlsxwriter.Workbook(u'zem/N1_%s' % self.sub + u'_Земля_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'N1_Земля')
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
	      
	       
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):
	       for x in range(1,int(self.pag)+1):
                    yield Task ('post',url=page+'?page=%d'%x,network_try_count=100)
	          
	       
            
	  def task_post(self,grab,task):
	       links = grab.doc.select(u'//a[@data-test="offers-list-item-header"]')
	       for elem in links:
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
     
	  def task_item(self, grab, task):
	       try:
		    ray = grab.doc.select(u'//a[@class="breadcrumbs-list__link"][contains(@href,"rayon")]').text()
	       except IndexError:
		    ray = ''          
	       try:
		    punkt = grab.doc.select(u'//div[@class="user-menu"]/ul[2]').text()
	       except IndexError:
		    punkt = ''
	       try:
		    ter= grab.doc.select(u'//p[contains(text(),"Ориентир:")]').text().replace(u'Ориентир: ','')
	       except IndexError:
		    ter =''
	       try:
		    uliza=''
	       except IndexError:
	            uliza = ''
	       try:
		    a2 = grab.doc.select(u'//h1').text().split(' — ')[1]
		    count2 = len(a2.split(','))-1
		    if count2 == 3:
			 dom= re.sub('[^\d]','',a2.split(', ')[2])[:2]
		    elif count2 == 2:
			 dom = re.sub('[^\d]','',a2.split(', ')[1])[:2]
		    elif count2 == 1:
			 dom = ''#re.sub('[^\d]','',a2.split(', ')[0])[:2]
		    else:
			 dom=''
	       except IndexError:
	            dom = ''
		     
	     
		 
	       try:
		    metro = grab.doc.select(u'//p[@id="priceMulti_3_0"]/strong[1]').text()
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//th[contains(text(),"Водоснабжение")]/following-sibling::td').text()
		 #print rayon
	       except IndexError:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//span[@class="object_item_metro_comment"]').text().split(u'мин. ')[1]
	       except IndexError:
		    metro_tr = ''
		    
	       
		    
	       
		   
	       try:
		    price = grab.doc.select(u'//span[@data-test="offer-card-price"]').text()+u' р.'
		 #print price + u' руб'	    
	       except IndexError:
		    price = ''
		   
	       
		     
	       
     
	       
     
	       try:
		    plosh_ob = grab.doc.select(u'//span[contains(text(),"Площадь участка")]/following-sibling::span').text()
		  #print rayon
	       except IndexError:
		    plosh_ob = ''
     
	       
		    
	       try:
		    et = grab.doc.select(u'//h2').text().replace(u'Параметры','')
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		    etagn = re.sub('[^\d\.]','',grab.doc.select(u'//p[contains(text(),"Обновлено")]').text().split(', ')[0])
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''
		     
	       
		
		     
	       try:
		    opis = grab.doc.select(u'//div[@class="offer-card-description__text"]').text()
		    
	       except IndexError:
	            opis = ''
		
	       try:
		    phone = re.sub('[^\d\+]','',grab.doc.select(u'//li[@class="offer-card-contacts-phones__item"]/a/@href').text())
	       except IndexError:
		    phone = ''
		   
	       try:
		    lico = grab.doc.select(u'//div[@class="offer-card-contacts__person _name"]').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//div[@class="offer-card-contacts__person _agency"]').text()
		 #print rayon
	       except IndexError:
		    comp = ''
		    
	       try:
		    data = re.sub('[^\d\.]','',grab.doc.select(u'//p[contains(text(),"Обновлено")]').text().split(', ')[1])
	       except IndexError:
		    data = ''
		    
	       
		    
	       
		   
		   
	       
		   
	      
	   
	       projects = {'sub': self.sub,
		           'rayon': ray,
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
		           'lico':lico.replace(comp,''),
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
	       print  task.project['etashost']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['etach']
	       
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 6,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 11,task.project['metro'])
	       self.ws.write(self.result, 16,task.project['naz'])
	       #self.ws.write(self.result, 10,task.project['object'])
	       self.ws.write(self.result, 9,oper)
	       self.ws.write(self.result, 10, task.project['cena'])
	       #self.ws.write(self.result, 13, task.project['cena_m'])
	       #self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 12, task.project['plosh_ob'])
	       #self.ws.write(self.result, 16, task.project['plosh_gil'])
	       #self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       #self.ws.write(self.result, 18, task.project['plosh_com'])
	       self.ws.write(self.result, 31, task.project['etach'])
	       self.ws.write(self.result, 29, task.project['etashost'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'N1.RU')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 28, task.project['data'])
	       self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	      
	       
	       print('*'*50)
	       print self.sub
	      
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*50)
	       self.result+= 1

	       
	       #if self.result >= 10:
		    #self.stop()

	 
     
     bot = Nndv_Zem(thread_number=3,network_try_limit=1000)
     bot.load_proxylist('/home/oleg/Proxy/tipa.txt','text_file')
     bot.create_grab_instance(timeout=500, connect_timeout=5000)
     bot.run()
     #print bot.sub,bot.end
     print('Спим 2 сек...')
     time.sleep(2)
     print('Сохранение...')
     bot.workbook.close()
     print('Done!')
     time.sleep(1)     
     i=i+1
     try:
          page = l[i]
     except IndexError:
          if oper == u'Продажа':
               i = 0
               l= open('Links/Zem_Arenda.txt').read().splitlines()
               page = l[i]
               oper = u'Аренда'
	  else:
               break

     
     

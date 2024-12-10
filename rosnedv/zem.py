#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import time
import math
import os
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



i = 0
l= open('links/zem.txt').read().splitlines()

page = l[i] 






while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Nedvizhka_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       #self.link =l[i]
	       for p in range(1,51):
		    try:
                         time.sleep(2)
			 g = Grab(timeout=5, connect_timeout=5)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                         g.go(self.f)
			 print g.doc.code,p,'/ 50'
	                 self.sub = g.doc.select(u'//div[@class="location-wrap"]/a').text()
			 print self.sub
	                 #self.num = re.sub('[^\d]','',g.doc.select(u'//p[@class="total_offers"]').text())
	                 #self.pag = int(math.ceil(float(int(self.num))/float(20)))
	                 #print self.sub,self.pag,self.num
			 del g
	                 break
		    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
		         print g.config['proxy'],'Change proxy'
		         g.change_proxy()
			 del g
		         continue
                    
	       else:
	            self.sub = ''
		    self.pag = 0
		    self.stop()
	       
	       self.workbook = xlsxwriter.Workbook(u'zem/Rosnedv_%s' % self.sub + u'_Земля_'+str(i+1)+'.xlsx')
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
	       self.ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 30, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):
	       for x in range(1,50):
	            yield Task ('post',url=self.f+'more_realty/?page=%d'%x,refresh_cache=True,network_try_count=100)
	          
	       
            
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[contains(@href,"uchastki")]'):
                    ur = 'https://www.rosnedv.ru'+elem.attr('href').split('"')[1].replace('\/','/')[:-1]                    
		    #if len(ur) > 56:
                    yield Task('item', url=ur,refresh_cache=True, network_try_count=100,valid_status=(500,501,502))
     
	  def task_item(self, grab, task):
	       if grab.doc.code == 200:
		    try:
			 ray = grab.doc.select(u'//div[contains(text(),"Район")]/following-sibling::div').text()
		    except IndexError:
			 ray = ''          
		    try:
			 punkt= grab.doc.select(u'//small').text().split(', ')[1]
		    except IndexError:
			 punkt = ''
		    try:
			 ter = grab.doc.select(u'//div[contains(text(),"Микрорайон")]/following-sibling::div').text()
		    except IndexError:
			 ter =''
		    try:
			 uliza = grab.doc.select(u'//small').text().split(', ')[2]
		    except (IndexError,UnboundLocalError):
			 uliza = ''
		    try:
			 #dm = grab.doc.select(u'//span[contains(text(),"Адрес")]/following::div[1]').text()
			 dom = grab.doc.select(u'//small').text().split(', ')[3]
		    except (IndexError,AttributeError):
			 dom = ''
			  
		    
		      
		    try:
			 metro = grab.doc.select(u'//th[contains(text(),"Газ:")]/following-sibling::td').text()
		    except IndexError:
			 metro = ''
			
		    try:
			 metro_min = grab.doc.select(u'//th[contains(text(),"Водоснабжение:")]/following-sibling::td').text()
		      #print rayon
		    except IndexError:
			 metro_min = ''
			
		    try:
			 metro_tr = grab.doc.select(u'//th[contains(text(),"Электричество:")]/following-sibling::td').text()
		    except IndexError:
			 metro_tr = ''
     
		    try:
			 price = grab.doc.select(u'//div[@class="price"]').text()
		      #print price + u' руб'	    
		    except IndexError:
			 price = ''
		    try:
			 try:
			      plosh_ob = grab.doc.select(u'//div[contains(text(),"Площадь")]/following-sibling::div').text()
			 except IndexError:
			      plosh_ob = grab.doc.select(u'//th[contains(text(),"Площадь общая:")]/following-sibling::td').text()
		    except IndexError:
			 plosh_ob = ''
	  
		    
			 
		    try:
			 et = grab.doc.select(u'//small').text()
		      #print price + u' руб'	    
		    except IndexError:
			 et = '' 
			
		    try:
			 etagn = grab.doc.select(u'//div[@class="info-status"]').text().split(u'Добавлено ')[1].split(u' Обновлено ')[0]
		    except IndexError:
			 etagn = ''
     
		    try:
			 opis = grab.doc.select(u'//div[@class="desc-wrap"]').text().replace(u'Описание от продавца ','') 
		    except IndexError:
			 opis = ''
		     
		    try:
			 phone = re.sub('[^\d\+]','',grab.doc.select(u'//a[@itemprop="telephone"]').text())
		    except IndexError:
			 phone = ''
			
		    try:
			 lico = grab.doc.select(u'//div[@class="spec-name"]').text()
		    except IndexError:
			 lico = ''
			 
		    try:
			 comp = grab.doc.select(u'//title').text().split(' ')[0]
		      #print rayon
		    except IndexError:
			 comp = ''
			 
		    try:
			 data = grab.doc.select(u'//div[@class="info-status"]').text().split(u' Обновлено ')[1]
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
			        'lico':lico,
			        'company':comp,
			        'data':data}
		  
		  
		  
		    yield Task('write',project=projects,grab=grab)
	     
	     
	     
	     
	     
	     
	  def task_write(self,grab,task):
	       
	       print('*'*50)	       
	       print  task.project['sub']
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
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['data']
	       print  task.project['etashost']
	       print  task.project['rayon']
	       
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 31,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 15,task.project['metro'])
	       self.ws.write(self.result, 16,task.project['naz'])
	       self.ws.write(self.result, 18,task.project['tran'])
	       self.ws.write(self.result, 10, task.project['cena'])
	       #self.ws.write(self.result, 13, task.project['cena_m'])
	       #self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 12, task.project['plosh_ob'])
	       #self.ws.write(self.result, 16, task.project['plosh_gil'])
	       #self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       #self.ws.write(self.result, 18, task.project['plosh_com'])
	       self.ws.write(self.result, 31, task.project['etach'])
	       self.ws.write(self.result, 28, task.project['etashost'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'Rosnedv.RU')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 9, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	      
	       
	       print('*'*50)
	       print 'Ready - '+str(self.result)
	       print '***',i+1,'/',len(l),'***'
	       print task.project['company']
	       print('*'*50)
	       self.result+= 1
	       
	       
	       
	       
	       #if self.result > 10:
		    #self.stop()

     bot = Nedvizhka_Zem(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     bot.workbook.close()
     print('Done')
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break          

time.sleep(5)
os.system("/home/oleg/pars/rosnedv/comm.py")
     

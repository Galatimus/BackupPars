#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import time
import os
import math
from datetime import datetime,timedelta
import random
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



i = 0
l= open('links/urls.txt').read().splitlines()

page = l[i] 
oper = u'Продажа'

while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Nedvizhka_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       for p in range(1,16):
		    try:
			 time.sleep(2)
			 g = Grab(timeout=25, connect_timeout=65)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f+'/prodazha_domov/')
			 self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="total"]/p').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 print('*'*50)
			 print self.num
			 print self.pag
			 print('*'*50)
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	  
	       else:
		    self.num = 1
		    self.pag = 1
		    self.stop()
                    
		    
	       
	       self.workbook = xlsxwriter.Workbook(u'zagg/Move_Загород_'+oper+str(i)+'.xlsx')
	       self.ws = self.workbook.add_worksheet()
	       self.ws.write(0, 0, "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	       self.ws.write(0, 1, "МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	       self.ws.write(0, 2, "НАСЕЛЕННЫЙ_ПУНКТ")
	       self.ws.write(0, 3, "ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	       self.ws.write(0, 4, "УЛИЦА")
	       self.ws.write(0, 5, "ДОМ")
	       self.ws.write(0, 6, "ОРИЕНТИР")
	       self.ws.write(0, 7, "ТРАССА")
	       self.ws.write(0, 8, "УДАЛЕННОСТЬ")
	       self.ws.write(0, 9, "КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
	       self.ws.write(0, 10, "ТИП_ОБЪЕКТА")
	       self.ws.write(0, 11, "ОПЕРАЦИЯ")
	       self.ws.write(0, 12, "СТОИМОСТЬ")
	       self.ws.write(0, 13, "ЦЕНА_М2")
	       self.ws.write(0, 14, "ПЛОЩАДЬ_ОБЩАЯ")
	       self.ws.write(0, 15, "КОЛИЧЕСТВО_КОМНАТ")
	       self.ws.write(0, 16, "ЭТАЖНОСТЬ")
	       self.ws.write(0, 17, "МАТЕРИАЛ_СТЕН")
	       self.ws.write(0, 18, "ГОД_ПОСТРОЙКИ")
	       self.ws.write(0, 19, "ПЛОЩАДЬ_УЧАСТКА")
	       self.ws.write(0, 20, "ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	       self.ws.write(0, 21, "ГАЗОСНАБЖЕНИЕ")
	       self.ws.write(0, 22, "ВОДОСНАБЖЕНИЕ")
	       self.ws.write(0, 23, "КАНАЛИЗАЦИЯ")
	       self.ws.write(0, 24, "ЭЛЕКТРОСНАБЖЕНИЕ")
	       self.ws.write(0, 25, "ТЕПЛОСНАБЖЕНИЕ")
	       self.ws.write(0, 26, "ЛЕС")
	       self.ws.write(0, 27, "ВОДОЕМ")
	       self.ws.write(0, 28, "БЕЗОПАСНОСТЬ")
	       self.ws.write(0, 29, "ОПИСАНИЕ")
	       self.ws.write(0, 30, "ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 31, "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 32, "ТЕЛЕФОН")
	       self.ws.write(0, 33, "КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 34, "КОМПАНИЯ")
	       self.ws.write(0, 35, "ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 36, "ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 37, "ДАТА_ПАРСИНГА")
	       self.ws.write(0, 38, "КАТЕГОРИЯ_ЗЕМЛИ")
	       self.ws.write(0, 39, "МЕСТОПОЛОЖЕНИЕ")
	       self.result= 1
            
              
    
	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'/prodazha_domov/?page=%d'%x,refresh_cache=True,network_try_count=100)
	          
	       
            
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@class="search-item__title-link search-item__item-link"]'):
	            ur = grab.make_url_absolute(elem.attr('href'))
	            #print ur
	            yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
     
	  def task_item(self, grab, task):
	       
	       try:
		    sub = grab.doc.select(u'//li[@class="top-menu__item"]/span').text()
	       except IndexError:
		    sub = ''	       
	       try:
		    ray = grab.doc.select(u'//div[@class="breadcrumbs"]/span/a[contains(@title, "район")]').text()
	       except IndexError:
		    ray = ''          
	       try:
		    try:
			 punkt = grab.doc.select(u'//div[@class="breadcrumbs"]/span/a[contains(@title, "в г")]').text()
		    except IndexError:
			 punkt = grab.doc.select(u'//div[@class="breadcrumbs"]/span/a[contains(@title, "поселок")]').text()
	       except IndexError:
			 punkt = ''	      

	       try:
                    ter = grab.doc.select(u'//div[contains(text(),"Количество этажей:")]/following-sibling::div').text()
               except IndexError:
	            ter = ''
	       
	       
	       try:
		    try:
			 try:
			      uliza = grab.doc.select(u'//div[@class="breadcrumbs"]/span/a[contains(@title, "улице")]').text()
			 except IndexError:
			      uliza = grab.doc.select(u'//span[@class="geo-block__geo-info_no-link"][contains(text(),"улица")]').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//span[@class="geo-block__geo-info_no-link"][contains(text(),"проезде")]').text()
	       except IndexError:
	            uliza =''
		    
		    
		    
	       try:
		    dom = re.sub('[^\d]','',grab.doc.select(u'//h1/span[@class="object-title_page-title_tail"]').text())
                    #dom = re.compile(r'[0-9]+$',re.S).search(dm).group(0)
	       except IndexError:
		    dom = ''
		     
	       try:
		    orentir = grab.doc.select(u'//div[contains(text(),"Количество комнат:")]/following-sibling::div').text()
	       except IndexError:
		    orentir = ''              
		 
	       try:
		    metro = grab.doc.select(u'//div[contains(text(),"цена за")]/following-sibling::div').text()#.split('/')[1]
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//div[contains(text(),"Тип объекта:")]/following-sibling::div').text()
		 #print rayon
	       except IndexError:
		    metro_min = ''
		    
	       try:
	            mat = grab.doc.select(u'//div[contains(text()," Тип здания:")]/following-sibling::div').text()
	       except IndexError:
		    mat = ''		    
		   
	       try:
		    metro_tr = grab.doc.select(u'//div[@class="object-place__address"]').text()
	       except IndexError:
		    metro_tr = ''

	       try:
		    price = grab.doc.select(u'//div[contains(text(),"Цена:")]/following-sibling::div').text()
		 #print price + u' руб'	    
	       except IndexError:
		    price = ''

	       try:
		    plosh_ob = grab.doc.select(u'//div[contains(text(),"Общая площадь:")]/following-sibling::div').text()
	       except IndexError:
		    plosh_ob = ''
		    
	       try:
	            elekt = grab.doc.select(u'//div[contains(text()," Площадь участка:")]/following-sibling::div').text()
	       except IndexError:
	            elekt =''               
	       
		    
	       try:
		    et = grab.doc.select(u'//div[contains(text(),"Водопровод:")]/following-sibling::div').text()
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
	       try:
		    opis = grab.doc.select(u'//h4[contains(text(),"Описание")]/following-sibling::div[1]').text() 
	       except IndexError:
		    opis = ''
		
	       try:
		    lico = grab.doc.select(u'//div[@class="block-user__name"]').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//div[@class="block-user__agency"]').text().replace(u'Риелтор','')
		 #print rayon
	       except IndexError:
		    comp = ''
		    
	   
	       projects = {'sub': sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
	                   'comnat': orentir,
		           'metro': metro,
	                   'material': mat,
	                   'naz': metro_min,		           
		           'tran': metro_tr,
		           'cena': price,		           
		           'plosh_ob':plosh_ob,		           
		           'etach': et,
	                   'pl_uh':elekt,
		           'opis':opis,
		           'url':task.url,
		           'lico':lico,
		           'company':comp}
		           
	       
	       try:
		    link = task.url.replace('objects','objectsv3/printing')
		    yield Task('phone',url=link,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    pass       
	     
	  def task_phone(self, grab, task):
	       try:
		    phone= re.sub('[^\d\,]','',grab.doc.select(u'//div[@class="phone"]').text())
	       except IndexError:
		    phone = ''
	       try:
		    data1=  grab.doc.select(u'//div[@class="tech-info"]/div[2]/span').text().split(' ')[1]
	       except IndexError:
		    data1 =''
	       try:
		    data = grab.doc.select(u'//div[@class="tech-info"]/div[1]/span').text().split(' ')[1]
	       except IndexError:
		    data = ''
	  
	       project2 ={'phone':phone,
			  'dataraz': data,
			  'dataob':data1}
	     
	     
	       yield Task('write',project=task.project,proj=project2,grab=grab)
	     
	     
	     
	     
	     
	     
	  def task_write(self,grab,task):
	       #if task.project['phone']<>'':
	       print('*'*50)	       
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['metro']
	       print  task.project['cena']
	       print  task.project['plosh_ob']
	       print  task.project['etach']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.proj['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.proj['dataraz']
               print  task.proj['dataob']
	       print  task.project['tran']
	       
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 16,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 13,task.project['metro'])
	       self.ws.write(self.result, 10,task.project['naz'])
	       self.ws.write(self.result, 39,task.project['tran'])
	       self.ws.write(self.result, 11,oper)
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 15, task.project['comnat'])
	       self.ws.write(self.result, 17, task.project['material'])
	       self.ws.write(self.result, 14, task.project['plosh_ob'])
	       self.ws.write(self.result, 19, task.project['pl_uh'])
	       #self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       #self.ws.write(self.result, 18, task.project['plosh_com'])
	       self.ws.write(self.result, 22, task.project['etach'])
	       self.ws.write(self.result, 35, task.proj['dataraz'])
	       self.ws.write(self.result, 29, task.project['opis'])
	       self.ws.write(self.result, 30, u'MOVE.RU')
	       self.ws.write_string(self.result, 31, task.project['url'])
	       self.ws.write(self.result, 32, task.proj['phone'])
	       self.ws.write(self.result, 33, task.project['lico'])
	       self.ws.write(self.result, 34, task.project['company'])
	       self.ws.write(self.result, 36, task.proj['dataob'])
	       self.ws.write(self.result, 37, datetime.today().strftime('%d.%m.%Y'))
	      
	       
	       print('*'*50)
	       print 'Ready - '+str(self.result)+'/'+self.num
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print  task.project['naz']
	       print('*'*50)
	       self.result+= 1
	       
	       if str(self.result) == str(self.num):
		    self.stop()	       

     bot = Nedvizhka_Zem(thread_number=10,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     time.sleep(2)
     bot.workbook.close()
     print('Done')   
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break          

     
     

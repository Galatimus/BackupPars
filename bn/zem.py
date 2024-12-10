 #!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import os
import time
import math
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logging.basicConfig(level=logging.DEBUG)





i = 0
l= open('links/zem_prod.txt').read().splitlines()

page = l[i] 
oper = u'Продажа'





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Mag_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       #self.link =l[i]
	       while True:
		    try:
                         time.sleep(2)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                         g.go(self.f)
                         self.sub = g.doc.select(u'//a[@id="region_btn"]').text().replace('/','-')
			 if g.doc.select(u'//div[@class="count"]').exists()==True:
			      self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="count"]').text())
			      self.pag = int(math.ceil(float(int(self.num))/float(50)))
			 else:
			      self.num = 0
			      self.pag = 0
			 print self.sub,self.pag,self.num
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
		         print g.config['proxy'],'Change proxy'
		         g.change_proxy()
			 del g
		         continue
                    
		    
	       
	       self.workbook = xlsxwriter.Workbook(u'zem/Bn_%s' % self.sub + u'_Земля_'+oper+str(i+1)+'.xlsx')
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
	       for x in range(0,self.pag+1):
                    yield Task ('post',url=self.f+'?start='+str(x*50),refresh_cache=True,network_try_count=100)
	          
	       
            
	  def task_post(self,grab,task):
	       links = grab.doc.select(u'//a[@class="underline"][@onclick="return false;"]')
	       for elem in links:
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
     
	  def task_item(self, grab, task):
	       try:
		    ray = grab.doc.select(u'//ul[@id="breadcrumb"]/li/a[contains(text(),"район")]').text()
	       except IndexError:
		    ray = ''          
	       try:
		    if self.sub == u"Москва":
                         punkt= u"Москва"
                    elif self.sub == u"Санкт-Петербург":
	                 punkt= u"Санкт-Петербург"
                    elif self.sub == u"Севастополь":
	                 punkt= u"Севастополь"
                    else:   
	                 punkt = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/text()').text().split(', ')[0]
	       except IndexError:
		    punkt = ''
	       #try:
		    #ter= grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Микрорайон")]/following-sibling::dd').text()
	       #except IndexError:
		    #ter =''
	       try:
		    #if self.sub == u"Москва":
                    uliza= grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/text()').text().split(', ')[1]
                    #else:
	                 #uliza= grab.doc.select(u'//h1').text().split(', ')[2]
	       except IndexError:
		    uliza = ''
	       try:
		    #if self.sub == u"Москва":
                    dom= grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/text()').text().split(', ')[2]
                    #else:
	                 #dom= grab.doc.select(u'//h1').text().split(', ')[3].replace(self.sub,'')
	       except IndexError:
		    dom = ''
		     
	       try:
		    orentir = grab.doc.select(u'//ul[@id="breadcrumb"]/li/a[contains(text(),"АО")]').text()
	       except IndexError:
		    orentir = ''              
		 
	       try:
		    metro = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/text()').text().replace(u'Объект на карте','')
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Удаленность от города")]/preceding-sibling::div').text()
		 #print rayon
	       except IndexError:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Категория земли")]/preceding-sibling::div').text()
	       except IndexError:
		    metro_tr = ''

	       try:
		    price = grab.doc.select(u'//dt[contains(text(),"Цена")]/following-sibling::dd').text()
		 #print price + u' руб'	    
	       except IndexError:
		    price = ''
	
	       try:
		    plosh_ob = grab.doc.select(u'//dt[contains(text(),"Площадь участка")]/following-sibling::dd').text()
	       except IndexError:
		    plosh_ob = ''

	       try:
		    et = grab.doc.select(u'//dt[contains(text(),"Статус участка")]/following-sibling::dd').text()
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		    etagn = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Газ")]/preceding-sibling::div').text()
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''
               try:
		    opis = grab.doc.select(u'//div[@id="description"]').text() 
	       except IndexError:
		    opis = ''
		
	       try:
		    phone = re.sub('[^\d\,]','',grab.doc.select(u'//dt[contains(text(),"Телефон")]/following-sibling::dd').text())
	       except IndexError:
		    phone = ''
		   
	       try:
		    lico = grab.doc.select(u'//dt[contains(text(),"Дата обновления")]/following-sibling::dd').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    try:
                         comp = grab.doc.select(u'//dt[contains(text(),"Продает")]/following-sibling::dd').text()
                    except IndexError:
	                 comp = grab.doc.select(u'//dt[contains(text(),"Сдает")]/following-sibling::dd').text()
	       except IndexError:
		    comp = ''
		    
	       try: 
	            data = grab.doc.select(u'//dt[contains(text(),"Дата размещения")]/following-sibling::dd').text()
	       except IndexError:
	            data=''

	   
	       projects = {'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor':orentir,
		           'ulica': uliza,
		           'dom': dom,
		           'metro': self.sub+', '+metro,
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
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
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
	       print  task.project['metro']
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 31,task.project['metro'])
	       self.ws.write(self.result, 8,task.project['naz'])
	       self.ws.write(self.result, 13,task.project['tran'])
	       self.ws.write(self.result, 9,oper)
	       self.ws.write(self.result, 10, task.project['cena'])
	       #self.ws.write(self.result, 13, task.project['cena_m'])
	       #self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 12, task.project['plosh_ob'])
	       self.ws.write(self.result, 14, task.project['etach'])
	       self.ws.write(self.result, 15, task.project['etashost'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'Система "Бюллетень Недвижимости"')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 29, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 28, task.project['data'])
	       self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	      
	       
	       print('*'*50)
	       print self.sub
	      
	       print 'Ready - '+str(self.result)+'/'+self.num
	       print 'Tasks - %s' % self.task_queue.size()
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*50)
	       self.result+= 1
	       
	       
	       
	       
	       #if self.result > 10:
		    #self.stop()

     
     bot = Mag_Zem(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...') 
     command = 'mount -a'
     os.system('echo %s|sudo -S %s' % ('1122', command))
     time.sleep(2)
     bot.workbook.close()
     print('Done')

     i=i+1
     try:
	  page = l[i]
     except IndexError:
	  if oper == u'Продажа':
	       i = 0
	       l= open('links/zem_arenda.txt').read().splitlines()
	       dc = len(l)
	       page = l[i]
	       oper = u'Аренда'
	  else:
	       break

     
time.sleep(5)
os.system("/home/oleg/pars/bn/comm.py")
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

g = Grab(timeout=30, connect_timeout=100)

i = 0
l= open('Links/Zag_Prod.txt').read().splitlines()

page = l[i] 
oper = u'Продажа'





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class N1_Zag(Spider):
	  def prepare(self):
	       self.f = page
	       #self.link =l[i]
	       while True:
		    try:
                         time.sleep(2)
                         g.go(self.f)
                         if g.doc.select(u'//a[@data-test="offers-list-item-header"]').exists()==True:
                              self.num = re.sub('[^\d]','',g.doc.select(u'//div[@data-test="offers-list-header-results"]').text())
                              self.pag = int(math.ceil(float(int(self.num))/float(25)))
                              self.sub = g.doc.select(u'//div[@class="user-menu"]/ul[2]').text()
                              print self.sub,self.pag,self.num
                              break
                         else:
	                      self.sub=''
	                      self.pag=1
	                      self.num=1
	                      break
                    except(GrabTimeoutError,GrabNetworkError, GrabConnectionError):
                         print g.config['proxy'],'Change proxy'
                         g.change_proxy()
                         continue
                    
		    
	       
	       self.workbook = xlsxwriter.Workbook(u'zag/N1_%s' % self.sub + u'_Загород_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Загород')
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
	       self.ws.write(0, 36, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 37, "ДАТА_ПАРСИНГА")
	       
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=page+'?page=%d'%x,network_try_count=100)   
	       
            
	  def task_post(self,grab,task):
	       links = grab.doc.select(u'//a[@data-test="offers-list-item-header"]')
	       for elem in links:
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100,use_proxylist=False)
	      
     
	  def task_item(self, grab, task):
	       try:
		    ray = re.sub('[^\d\.]','',grab.doc.select(u'//p[contains(text(),"Обновлено")]').text().split(', ')[0])
	       except IndexError:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//div[@class="user-menu"]/ul[2]').text()
	       except IndexError:
		    punkt = ''
	       try:
		    try:
                         ter= grab.doc.select(u'//a[@class="breadcrumbs-list__link"][contains(@href,"district")]').text()
                    except IndexError:
	                 ter= grab.doc.select(u'//a[@class="breadcrumbs-list__link"][contains(@href,"rayon")]').text()
	       except IndexError:
		    ter =''
	       try:
		    uliza = grab.doc.select(u'//a[@class="breadcrumbs-list__link"][contains(@href,"street")]').text()
	       except IndexError:
		    uliza = ''
	       try:
		    dom=grab.doc.rex_text(u'house_number":"(.+?)"').replace('",','')
	       except IndexError:
		    dom = ''
		     
	       try:
		    orentir = grab.doc.select(u'//a[@class="breadcrumbs-list__link"][contains(@href,"microdistrict")]').text()
	       except IndexError:
		    orentir = ''              
		 
	       try:
		    metro = grab.doc.select(u'//a[@class="breadcrumbs-list__link"][contains(@href,"type")]').text()
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//li[@class="card-living-content-deal-params__item"][1]').text()
		 #print rayon
	       except IndexError:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//span[contains(text(),"Комнат")]/following-sibling::span').text()
	       except IndexError:
		    metro_tr = ''
	       try:
		    try:
		         price = grab.doc.select(u'//div[@class="card-living-content-price"]').text()+u' руб.'
		    except IndexError:
			 price = grab.doc.select(u'//div[@class="card-dacha-content-price"]').text()+u' руб.'
	       except IndexError:
		    price = '' 
               try:
                    tip = grab.doc.select(u'//span[contains(text(),"Материал дома")]/following-sibling::span').text()
               except IndexError:
                    tip = ''	       
	       try:
		    try:
		         plosh_ob = grab.doc.select(u'//span[contains(text(),"Общая площадь")]/following-sibling::span').text()
		    except IndexError:
			 plosh_ob = grab.doc.select(u'//span[contains(text(),"Площадь дома")]/following-sibling::span').text()
	       except IndexError:
		    plosh_ob = ''

	       try:
		    et = grab.doc.select(u'//span[contains(text(),"Этажей")]/following-sibling::span').text()
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		    etagn = grab.doc.select(u'//span[contains(text(),"Площадь участка")]/following-sibling::span').text()
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''

	       try:
		    opis = grab.doc.select(u'//div[@class="offer-card-description__text"]').text() 
	       except IndexError:
		    opis = ''
		
	       try:
		    phone = re.sub('[^\d]','',grab.doc.rex_text(u'formatted(.*?)comment'))
	       #print phone
	       except (IndexError,GrabConnectionError,GrabNetworkError,GrabTimeoutError,IOError):
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
	                   'plosh':orentir,
		           'etach': et,
		           'etashost': etagn,      
		           'opis':opis,
		           'url':task.url,
		           'phone':phone,
		           'lico':lico,
	                   'tip':tip,
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
	       print  task.project['plosh']	       
	       print  task.project['metro']
	       print  task.project['naz']	      
	       print  task.project['tran']
	       print  task.project['cena']	       
	       print  task.project['plosh_ob']
	       print  task.project['etach']
	       print  task.project['etashost']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['tip']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['rayon']
	       
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 36,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 10,task.project['metro'])
	       self.ws.write(self.result, 13,task.project['naz'])
	       self.ws.write(self.result, 17,task.project['tip'])
	       self.ws.write(self.result, 11,oper)
	       self.ws.write(self.result, 12, task.project['cena'])
	       #self.ws.write(self.result, 13, task.project['cena_m'])
	       self.ws.write(self.result, 6, task.project['plosh'])
	       self.ws.write(self.result, 14, task.project['plosh_ob'])
	       #self.ws.write(self.result, 16, task.project['plosh_gil'])
	       #self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       self.ws.write(self.result, 15, task.project['tran'])
	       self.ws.write(self.result, 16, task.project['etach'])
	       self.ws.write(self.result, 19, task.project['etashost'])
	       self.ws.write(self.result, 29, task.project['opis'])
	       self.ws.write(self.result, 30, u'N1.RU_Недвижимость')
	       self.ws.write_string(self.result, 31, task.project['url'])
	       self.ws.write(self.result, 32, task.project['phone'])
	       self.ws.write(self.result, 33, task.project['lico'])
	       self.ws.write(self.result, 34, task.project['company'])
	       self.ws.write(self.result, 35, task.project['data'])
	       self.ws.write(self.result, 37, datetime.today().strftime('%d.%m.%Y'))
	      
	       
	       print('*'*50)
	       print self.sub
	      
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*50)
	       self.result+= 1
	       
	       
	       
	       
	       #if self.result > 50:
		    #self.stop()
	

     
     bot = N1_Zag(thread_number=5,network_try_limit=2000)
     bot.load_proxylist('/home/oleg/Proxy/tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=500)
     bot.run()
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
               l= open('Links/Zag_Arenda.txt').read().splitlines()
               dc = len(l)
               page = l[i]
               oper = u'Аренда'
          else:
               break

     
     

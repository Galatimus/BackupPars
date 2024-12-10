#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import math
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
import os
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






i = 0
l= open('Links/kvv.txt').read().splitlines()
page = l[i]





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class Nedvizhka_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	      
               for p in range(1,51):
                    try:
                         time.sleep(2)
			 g = Grab(timeout=5, connect_timeout=10)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http')			 
                         g.go(self.f)
			
			 if g.doc.select(u'//div[@class="navigation"]').exists()==True:
			      print g.doc.code
			      self.num = re.sub('[^\d]','',g.doc.select(u'//span[@class="search-count"]').text())
			      self.pag = str(float(math.ceil(float(int(self.num))/float(20)))).replace('.0','')
			      self.sub = g.doc.select(u'//li[@class="has-child"]/a').text()
			      print self.sub,self.pag,self.num
			      del g
			      break
			 else:
			      print 'Ждемс...'
			      time.sleep(60)
			      del g
			      continue
			 
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.sub = ''
		    self.pag = 0
		    self.num=1
		    
		    
	       self.workbook = xlsxwriter.Workbook(u'flats/Nedvizhka_Жилье_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet()
	       self.ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	       self.ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	       self.ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
	       self.ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	       self.ws.write(0, 4, u"УЛИЦА")
	       self.ws.write(0, 5, u"ДОМ")
	       self.ws.write(0, 6, u"ОРИЕНТИР")
	       self.ws.write(0, 7, u"СТАНЦИЯ_МЕТРО")
	       self.ws.write(0, 8, u"ДО_МЕТРО_МИНУТ")
	       self.ws.write(0, 9, u"ПЕШКОМ_ТРАНСПОРТОМ")
	       self.ws.write(0, 10, u"ТИП_ОБЪЕКТА")
	       self.ws.write(0, 11, u"ОПЕРАЦИЯ")
	       self.ws.write(0, 12, u"СТОИМОСТЬ")
	       self.ws.write(0, 13, u"ЦЕНА_М2")
	       self.ws.write(0, 14, u"КОЛИЧЕСТВО_КОМНАТ")
	       self.ws.write(0, 15, u"ПЛОЩАДЬ_ОБЩАЯ")
	       self.ws.write(0, 16, u"ПЛОЩАДЬ_ЖИЛАЯ")
	       self.ws.write(0, 17, u"ПЛОЩАДЬ_КУХНИ")
	       self.ws.write(0, 18, u"ПЛОЩАДЬ_КОМНАТ")
	       self.ws.write(0, 19, u"ЭТАЖ")
	       self.ws.write(0, 20, u"ЭТАЖНОСТЬ")
	       self.ws.write(0, 21, u"МАТЕРИАЛ_СТЕН")
	       self.ws.write(0, 22, u"ГОД_ПОСТРОЙКИ")
	       self.ws.write(0, 23, u"РАСПОЛОЖЕНИЕ_КОМНАТ")
	       self.ws.write(0, 24, u"БАЛКОН")
	       self.ws.write(0, 25, u"ЛОДЖИЯ")
	       self.ws.write(0, 26, u"САНУЗЕЛ")
	       self.ws.write(0, 27, u"ОКНА")
	       self.ws.write(0, 28, u"СОСТОЯНИЕ")
	       self.ws.write(0, 29, u"ВЫСОТА_ПОТОЛКОВ")
	       self.ws.write(0, 30, u"ЛИФТ")
	       self.ws.write(0, 31, u"РЫНОК")
	       self.ws.write(0, 32, u"КОНСЬЕРЖ")
	       self.ws.write(0, 33, u"ОПИСАНИЕ")
	       self.ws.write(0, 34, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 35, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 36, u"ТЕЛЕФОН")
	       self.ws.write(0, 37, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 38, u"КОМПАНИЯ")
	       self.ws.write(0, 39, u"ДАТА_РАЗМЕЩЕНИЯ_ОБЪЯВЛЕНИЯ")
	       self.ws.write(0, 40, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 41, u"МЕСТОПОЛОЖЕНИЕ")
	       self.result= 1
	       
                
                
                
                
	  def task_generator(self):
	       for x in range(1,int(self.pag)+1):
                    yield Task ('post',url=self.f+'?page=%d'%x+'&grid_type=table',refresh_cache=True,network_try_count=20)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//tr[@class="property"]/td[3]/a'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=20)
	      
            
	 
        
	  def task_item(self, grab, task):
	       try:
		    if self.sub == u'Москва и область':
			 mesto = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Микрорайон")]/following-sibling::dd').text()
		    else:
			 mesto = grab.doc.select(u'//header[@class="property-title"]/figure/a[2][contains(text()," район")]').text()
	       except IndexError:
	            mesto =''
		    
	       try:
	            punkt = grab.doc.select(u'//header[@class="property-title"]/figure/a[1]').text().replace(self.sub,'')
	       except IndexError:
	            punkt = ''	       
		
               try:
		    if grab.doc.select(u'//header[@class="property-title"]/figure/a[2][contains(text()," район")]').exists() == False:
			 ter =  grab.doc.select(u'//header[@class="property-title"]/figure/a[2]').text()
		    else:
			 ter =''
               except IndexError:
                    ter =''
               try:
                    uliza = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Улица")]/following-sibling::dd').text()
               except IndexError:
                    uliza = ''
               try:
                    dom = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Номер дома")]/following-sibling::dd').number()
               except IndexError:
                    dom = ''
		    
               try:
                    tip = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Тип")]/following-sibling::dd').text()
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Этаж")]/following-sibling::dd[1]').text()
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Этажность")]/following-sibling::dd').text()
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//dt[contains(text(),"Цена")]/following-sibling::dd[1]').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Площадь общая")]/following-sibling::dd[1]').text()+u' м2'
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Материал стен")]/following-sibling::dd').text()
               except IndexError:
                    ohrana =''
               try:
                    gaz =  grab.doc.rex_text(u"var rname=(.*?);").replace("'","")
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Метро")]/following-sibling::dd').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Комнат")]/following-sibling::dd/strong').number()
               except IndexError:
                    kanal =''
               try:
                    elek = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Площадь жилая")]/following-sibling::dd[1]').text()+u' м2'
               except IndexError:
                    elek =''
               try:
                    teplo = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Площадь кухни")]/following-sibling::dd[1]').text()+u' м2'
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//p[@itemprop="description"]').text() 
	       except IndexError:
	            opis = ''
               try:
		    try:
                         lico = grab.doc.select(u'//div[@class="agent-contact-info"]/div/h3').text()
		    except IndexError: 
			 lico = grab.doc.select(u'//dt[contains(text(),"Агент")]/following-sibling::dd[1]').text()
               except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//dd[contains(text(),"Организация")]/following-sibling::dt[1]').text()
               except IndexError:
                    comp = ''
               try:
                    data1 = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Год постройки")]/following-sibling::dd').number() 
               except IndexError:   
                    data1 = ''
	       try: 
	            data = grab.doc.select(u'//dt[contains(text(),"Дата подачи")]/following-sibling::dd[1]').text()
	       except IndexError:
		    data=''
	       
	       try:
                    phone = grab.doc.select(u'//dt[contains(text(),"Телефон")]/following-sibling::dd').text()
               except IndexError:
	            phone = ''
          
	       try:
		    if 'prodazha' in task.url:
			 oper = u'Продажа' 
		    elif 'arenda' in task.url:
			 oper = u'Аренда'
		    else:
			 oper = ''
	       except IndexError:
	            oper = ''	       
	       clearText = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", opis)
	       clearText = re.sub(u"[.,\-\s]{3,}", " ", clearText)
     
	
               projects = {'sub': self.sub,
	                  'adress': mesto,
	                   'terit':ter, 
	                   'punkt':punkt, 
	                   'ulica':uliza,
	                   'dom':dom,
	                   'tip':tip,
	                   'naz':naz.replace(klass,''),
	                   'klass': klass,
	                   'cena': price,
	                  'operacia': oper,
	                   'plosh': plosh,
	                   'ohrana':ohrana,
	                   'gaz': gaz,
	                   'voda': voda,
	                   'kanaliz': kanal,
	                   'electr': elek,
	                   'teplo': teplo,
	                   'opis': clearText,
	                   'url': task.url,
	                   'phone': phone,
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
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['data1']
	  
	       
	       
     
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['adress'])
	       self.ws.write(self.result, 3, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 10, task.project['tip'])
	       self.ws.write(self.result, 19, task.project['naz'])
	       self.ws.write(self.result, 20, task.project['klass'])
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 15, task.project['plosh'])
	       self.ws.write(self.result, 21, task.project['ohrana'])
	       self.ws.write(self.result, 41, task.project['gaz'])
	       self.ws.write(self.result, 7, task.project['voda'])
	       self.ws.write(self.result, 14, task.project['kanaliz'])
	       self.ws.write(self.result, 16, task.project['electr'])
	       self.ws.write(self.result, 17, task.project['teplo'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, u'Nedvizhka.RU')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 22, task.project['data1'])
	       self.ws.write(self.result, 40, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 11, task.project['operacia'])
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+self.num
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print '***',i+1,'/',len(l),'***'
	       print task.project['operacia'] 
	       print('*'*100)
	       self.result+= 1
	       
	      
	       
	       
	       
	       #if self.result > 10:
	            #self.stop()	       
	
	       #if int(self.result) >= int(self.num)-1:
	            #self.stop()		    	


     bot = Nedvizhka_Com(thread_number=7, network_try_limit=200)
     bot.load_proxylist('../tipa.txt','text_file')
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
     bot.workbook.close()
     print('Done') 
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
       
     
     
     
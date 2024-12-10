#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import math
import re
import time
import os
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)






i = 0
l= open('links/com.txt').read().splitlines()
page = l[i]





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class move_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	      
               for p in range(1,51):
                    try:
                         time.sleep(2)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                         g.go(self.f)
			 print g.doc.code,p,'/ 50'
			 self.sub = g.doc.select(u'//div[@class="location-wrap"]/a').text()
			 print self.sub
			 #self.num = g.doc.select(u'//a[contains(text(),"Показать")]/span').text()
			 #print self.num
			 #self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 #print self.sub,self.pag,self.num
			 del g
			 break
			 
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError,ValueError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
	            self.sub = ''
	            self.pag = 0
		    self.stop()	       
		    
	       self.workbook = xlsxwriter.Workbook(u'com/Rosnedv_%s' % bot.sub + '_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'rosnedv_Коммерческая')
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
	       self.result= 1
	      
	       
                
                
                
                
	  def task_generator(self):
	       for x in range(1,200):
                    yield Task ('post',url=self.f+'more_realty/?page=%d'%x,refresh_cache=True,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[contains(@href,"kommercheskaya_nedvijimost")]'):
		    ur = 'https://www.rosnedv.ru'+elem.attr('href').split('"')[1].replace('\/','/')[:-1]
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100,valid_status=(500,501,502))
	      
	            
	 
        
	  def task_item(self, grab, task):
	       
	       if grab.doc.code == 200:	       
		    try:
			 mesto = grab.doc.select(u'//small').text()
		    except IndexError:
			 mesto =''
		    try:
			 mesto1 = grab.doc.select(u'//div[contains(text(),"Район")]/following-sibling::div').text()
		    except IndexError:
			 mesto1 =''		    
		    try:
			 punkt = grab.doc.select(u'//small').text().split(', ')[1]
		    except IndexError:
			 punkt = ''	       
		     
		    try:
			 ter= grab.doc.select(u'//div[contains(text(),"Микрорайон")]/following-sibling::div').text()
		    except IndexError:
			 ter =''
		    try:
			 uliza= grab.doc.select(u'//small').text().split(', ')[2]
		    except (IndexError,UnboundLocalError):
			 uliza =''
		    try:
			 dom = grab.doc.select(u'//small').text().split(', ')[3]
			 #dom = re.compile(r'[0-9]+$',re.S).search(dm).group(0)
		    except (IndexError,AttributeError):
			 dom = ''
		      
		    try:
			 tip = grab.doc.select(u'//td[@class="param param-ext"][contains(text(),"Готовый ")]').text()#.split(' - ')[1]
		    except IndexError:
			 tip = ''
		    try:
			 naz = grab.doc.select(u'//div[contains(text(),"Тип")]/following-sibling::div').text()
		    except IndexError:
			 naz =''
		    try:
			 klass =  grab.doc.select(u'//title').text().split(' ')[0]
		    except IndexError:
			 klass = ''
		    try:
			 #try:
			 price = grab.doc.select(u'//div[@class="price"]').text()
			 #except IndexError:
			      #price = grab.doc.select(u'//span[@id="price-total-0"]').text()
		    except IndexError:
			 price =''
		    try: 
			 try:
			      plosh = grab.doc.select(u'//div[contains(text(),"Площадь")]/following-sibling::div').text()
			 except IndexError:
			      plosh = grab.doc.select(u'//th[contains(text(),"Площадь общая:")]/following-sibling::td').text()
		    except IndexError:
			 plosh=''
		    try:
			 ohrana = grab.doc.select(u'//th[contains(text(),"Метро:")]/following-sibling::td').text()
		    except IndexError:
			 ohrana =''
		    try:
			 gaz =  grab.doc.select(u'//title').text()
		    except IndexError:
			 gaz =''
		    try:
			 voda =  grab.doc.rex_text(u'Кадастровый номер: (.*?).')
		    except IndexError:
			 voda =''
		    try:
			 kanal = grab.doc.rex_text(u'"lng":(.*?),')
		    except IndexError:
			 kanal =''
		    try:
			 elek = grab.doc.rex_text(u'"lat":(.*?)}')
		    except IndexError:
			 elek =''
		    try:
			 teplo = grab.doc.select(u'//div[contains(text(),"Водоснабжение")]/following-sibling::div').text()
		    except IndexError:
			 teplo =''
		    #time.sleep(1)
		    try:
			 opis = grab.doc.select(u'//div[@class="desc-wrap"]').text().replace(u'Описание от продавца ','') 
		    except IndexError:
			 opis = ''
		    try:
			 lico = grab.doc.select(u'//div[@class="spec-name"]').text()
		    except IndexError:
			 lico = ''
		    try:
			 comp = grab.doc.select(u'//div[@class="company-block"]/p/a').text()
		    except IndexError:
			 comp = ''
		    
		    try:    
			 data = grab.doc.select(u'//div[@class="info-status"]').text().split(u' Обновлено ')[1]
		    except IndexError:
			 data =''
		    try:
			 data1 = grab.doc.select(u'//div[@class="info-status"]').text().split(u'Добавлено ')[1].split(u' Обновлено ')[0]
		    except IndexError:
			 data1=''
		    
		    try:
			 phone = re.sub('[^\d\+]','',grab.doc.select(u'//a[@itemprop="telephone"]').text())
		    except IndexError:
			 phone = ''
	       
		    
			 
	  
	     
		    projects = {'sub': self.sub,
			       'adress': mesto,
			       'adress1': mesto1,
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
	       print  task.project['adress1']
	       print  task.project['terit']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['tip']
	       print  task.project['naz']
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
	       print  task.project['adress']
	      
	  
	       
	       
     
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['adress1'])
	       self.ws.write(self.result, 3, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 7, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       self.ws.write(self.result, 28, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 26, task.project['ohrana'])
	       self.ws.write(self.result, 33, task.project['gaz'])
	       self.ws.write(self.result, 32, task.project['voda'])
	       self.ws.write(self.result, 35, task.project['kanaliz'])
	       self.ws.write(self.result, 34, task.project['electr'])
	       #self.ws.write(self.result, 21, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'Rosnedv.RU')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       #self.ws.write(self.result, 30, task.project['company'])
	       self.ws.write(self.result, 30, task.project['data'])
	       self.ws.write(self.result, 29, task.project['data1'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       #self.ws.write(self.result, 34, oper)
	       self.ws.write(self.result, 24, task.project['adress'])
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)
	       print '***',i+1,'/',len(l),'***'
	       print task.project['klass']
	       print('*'*100)
	       self.result+= 1
	       
	       #if self.result >= 50:
	            #self.stop()	       


     bot = move_Com(thread_number=5, network_try_limit=1000)
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
       
     
     
     
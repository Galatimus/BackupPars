#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import math
import os
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)






i = 0
l= open('Links/Com_Prod.txt').read().splitlines()
page = l[i]





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class Nndv_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]	      
               for p in range(1,21):
                    try:
                         time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')			 
                         g.go(self.f)
			 if 'kupit' in self.f:
			      self.oper = u'Продажа' 
			 elif 'snyat' in self.f:
			      self.oper = u'Аренда'
			 self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="breadcrumbs"]/ul/li/span[contains(text(),"объявлени")]').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(25)))
			 self.sub = g.doc.select(u'//span[@class="search-2gen-geo-filter-caption-link__text"]').text()
			 print self.sub,self.oper,self.pag,self.num
			 del g
			 break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.sub = ''
		    self.pag=1
		    self.num=1
		    
	       self.workbook = xlsxwriter.Workbook(u'com/N1_%s' % bot.sub + u'_Коммерческая_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet()
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
	       for x in range(1,self.pag+1):
		    yield Task ('post',url=page+'?page=%d'%x,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@class="link"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	       #yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
            
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//a[@data-test="offers-list-next-page"]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print 'no_page'	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//li[@class="geo-tags__item"]/a[contains(@href,"rayon")]/span[1]').text()
	       except IndexError:
	            mesto =''
		    
	       try:
	            punkt = grab.doc.select(u'//title').text().split('N1.RU ')[1]
	       except IndexError:
	            punkt = ''	       
		
               try:
                    ter =  grab.doc.select(u'//li[@class="geo-tags__item"]/a[contains(@href,"microdistrict")]/span[1]').text()
               except IndexError:
                    ter =''
               try:
                    uliza = grab.doc.select(u'//li[@class="geo-tags__item"]/a[contains(@href,"street")]/span[1]').text()
               except IndexError:
                    uliza = ''
               try:
                    dom =  grab.doc.select(u'//li[@class="breadcrumbs-list__item"]/a[contains(@href,"type")]/span[1]').text()
               except IndexError:
                    dom = ''
		    
               try:
                    tip = grab.doc.select(u'//li[@class="breadcrumbs-list__item"]/a[contains(@href,"purpose")]/span[1]').text()
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//div[@class="offer-card-content-location__reference"]').text()
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//span[contains(text(),"Этаж")]/following-sibling::span').text()
               except IndexError:
                    klass = ''
		    
	       try:
	            et = grab.doc.select(u'//span[contains(text(),"жность")]/following-sibling::span').text()
	       except IndexError:
	            et = ''		    
               try:
                    price = grab.doc.select(u'//span[@data-test="offer-card-price"]').text()+u' р.'
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//span[contains(text(),"Общая площадь")]/following-sibling::span').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//span[contains(text(),"Год постройки")]/following-sibling::span').text()
               except IndexError:
                    ohrana =''
               try:
                    gaz =  re.sub('[^\d\.]','',grab.doc.select(u'//p[contains(text(),"Обновлено")]').text().split(', ')[0])
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//h1').text()
               except IndexError:
                    voda =''
               try:
                    kanal =  grab.doc.rex_text(u'latitude":(.*?),')
               except IndexError:
                    kanal =''
               try:
                    elek = grab.doc.rex_text(u'longitude":(.*?),')
               except IndexError:
                    elek =''
               try:
		    ln = []
		    for m in grab.doc.select('//ul[@class="geo-tags__list"]/li'):
			 mes = m.text() 
			 ln.append(mes)
                    teplo = ', '.join(ln)
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//div[@class="foldable-description card-commercial-content__description"]').text() 
	       except IndexError:
	            opis = ''
               
               try:
                    comp = grab.doc.select(u'//div[@class="offer-card-contacts__person"]/a[contains(@href,"an")]/span').text()
               except IndexError:
                    comp = ''
	       try:
	            lico = grab.doc.select(u'//a[contains(@href,"users")]/span').text()
	       except IndexError:
	            lico = ''   
               
	       try: 
	            data = re.sub('[^\d\.]','',grab.doc.select(u'//p[contains(text(),"Обновлено")]').text().split(', ')[1])
	       except IndexError:
		    data=''
	       
	       try:
		    phone = re.sub('[^\d\+]','',grab.doc.select(u'//li[@class="offer-card-contacts-phones__item"]/a/@href').text())
               except IndexError:
	            phone = ''
          
	       
		    
     
	
               projects = {'sub': self.sub,
	                  'adress': mesto,
	                   'terit':ter, 
	                   'punkt':punkt, 
	                   'ulica':uliza,
	                   'dom':dom,
	                   'tip':tip,
	                   'naz':naz,
	                   'klass': klass.replace(et,''),
	                   'cena': price,
	                   'plosh': plosh,
	                   'ohrana':ohrana,
	                   'gaz': gaz,
	                   'ets': et,
	                   'voda': voda,
	                   'kanaliz': kanal,
	                   'electr': elek,
	                   'teplo': teplo+' '+naz,
	                   'opis': opis,
	                   'url': task.url,
	                   'phone': phone,
	                   'lico':lico,
	                   'company': comp,
	                   'data':data}
	                   
	  
	  
	       yield Task('write',project=projects,grab=grab)
	  
	  
	  
	  def task_write(self,grab,task):
	       if task.project['cena'] <> '':
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
		    print  task.project['ets']
		    print  task.project['cena']
		    print  task.project['plosh']
		    print  task.project['gaz']
		    print  task.project['voda']
		    print  task.project['kanaliz']
		    print  task.project['electr']
		    print  task.project['ohrana']
		    print  task.project['opis']
		    print  task.project['url']
		    print  task.project['phone']
		    print  task.project['lico']
		    print  task.project['company']
		    print  task.project['data']
		    print  task.project['teplo']
		   
	       
		    
		    
	  
		    self.ws.write(self.result, 0, task.project['sub'])
		    self.ws.write(self.result, 1, task.project['adress'])
		    self.ws.write(self.result, 3, task.project['terit'])
		    self.ws.write(self.result, 2, task.project['punkt'])
		    self.ws.write(self.result, 4, task.project['ulica'])
		    self.ws.write(self.result, 8, task.project['dom'])
		    self.ws.write(self.result, 9, task.project['tip'])
		    self.ws.write(self.result, 6, task.project['naz'])
		    self.ws.write(self.result, 15, task.project['klass'])
		    self.ws.write(self.result, 16, task.project['ets'])
		    self.ws.write(self.result, 11, task.project['cena'])
		    self.ws.write(self.result, 14, task.project['plosh'])
		    self.ws.write(self.result, 17, task.project['ohrana'])
		    self.ws.write(self.result, 30, task.project['gaz'])
		    self.ws.write(self.result, 33, task.project['voda'])
		    self.ws.write(self.result, 34, task.project['kanaliz'])
		    self.ws.write(self.result, 35, task.project['electr'])
		    self.ws.write(self.result, 24, task.project['teplo'])
		    self.ws.write(self.result, 18, task.project['opis'])
		    self.ws.write(self.result, 19, u'N1.RU')
		    self.ws.write_string(self.result, 20, task.project['url'])
		    self.ws.write(self.result, 21, task.project['phone'])
		    self.ws.write(self.result, 22, task.project['lico'])
		    self.ws.write(self.result, 23, task.project['company'])
		    self.ws.write(self.result, 29, task.project['data'])
		    self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
		    self.ws.write(self.result, 28, self.oper)
		    print('*'*100)
		    #print self.sub
		    print 'Ready - '+str(self.result)+'/'+str(self.num)
		    print 'Tasks - %s' % self.task_queue.size() 
		    print '***',i+1,'/',len(l),'***'
		    print self.oper
		    print('*'*100)
		    self.result+= 1
		    
		   
		    
		    
		    
		    #if self.result > 20:
			 #self.stop()	       


     bot = Nndv_Com(thread_number=5, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     bot.workbook.close()
     print('Done!')
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
       
     
     
     
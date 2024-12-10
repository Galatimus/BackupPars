#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import math
import re
import os
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)






i = 0
l= open('Links/flats_all.txt').read().splitlines()
page = l[i]


while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class Mag_Com(Spider):
	  def prepare(self):
	       self.f = page
               for p in range(1,51):
                    try:
                         time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http') 
                         g.go(self.f)
			 #self.num = re.sub('[^\d\,]','',g.doc.select(u'//title').text()).split(',')[0]
                         self.sub = g.doc.select(u'//span[contains(text(),"Регион:")]/following-sibling::span').text()#.replace(u'недвижимость в ','').replace(u'Москве',u'Москва')
                         print self.sub
			 #print self.num
			 del g
                         break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.sub = ''
		    #self.num = 1
		
			 
	       self.workbook = xlsxwriter.Workbook(u'flats/Realtymag_Жилье_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet()
	       self.ws.write(0, 0, "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	       self.ws.write(0, 1, "МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	       self.ws.write(0, 2, "НАСЕЛЕННЫЙ_ПУНКТ")
	       self.ws.write(0, 3, "ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	       self.ws.write(0, 4, "УЛИЦА")
	       self.ws.write(0, 5, "ДОМ")
	       self.ws.write(0, 6, "ОРИЕНТИР")
	       self.ws.write(0, 7, "СТАНЦИЯ_МЕТРО")
	       self.ws.write(0, 8, "ДО_МЕТРО_МИНУТ")
	       self.ws.write(0, 9, "ПЕШКОМ_ТРАНСПОРТОМ")
	       self.ws.write(0, 10, "ТИП_ОБЪЕКТА")
	       self.ws.write(0, 11, "ОПЕРАЦИЯ")
	       self.ws.write(0, 12, "СТОИМОСТЬ")
	       self.ws.write(0, 13, "ЦЕНА_М2")
	       self.ws.write(0, 14, "КОЛИЧЕСТВО_КОМНАТ")
	       self.ws.write(0, 15, "ПЛОЩАДЬ_ОБЩАЯ")
	       self.ws.write(0, 16, "ПЛОЩАДЬ_ЖИЛАЯ")
	       self.ws.write(0, 17, "ПЛОЩАДЬ_КУХНИ")
	       self.ws.write(0, 18, "ПЛОЩАДЬ_КОМНАТ")
	       self.ws.write(0, 19, "ЭТАЖ")
	       self.ws.write(0, 20, "ЭТАЖНОСТЬ")
	       self.ws.write(0, 21, "МАТЕРИАЛ_СТЕН")
	       self.ws.write(0, 22, "ГОД_ПОСТРОЙКИ")
	       self.ws.write(0, 23, "РАСПОЛОЖЕНИЕ_КОМНАТ")
	       self.ws.write(0, 24, "БАЛКОН")
	       self.ws.write(0, 25, "ЛОДЖИЯ")
	       self.ws.write(0, 26, "САНУЗЕЛ")
	       self.ws.write(0, 27, "ОКНА")
	       self.ws.write(0, 28, "СОСТОЯНИЕ")
	       self.ws.write(0, 29, "ВЫСОТА_ПОТОЛКОВ")
	       self.ws.write(0, 30, "ЛИФТ")
	       self.ws.write(0, 31, "РЫНОК")
	       self.ws.write(0, 32, "КОНСЬЕРЖ")
	       self.ws.write(0, 33, "ОПИСАНИЕ")
	       self.ws.write(0, 34, "ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 35, "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 36, "ТЕЛЕФОН")
	       self.ws.write(0, 37, "КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 38, "КОМПАНИЯ")
	       self.ws.write(0, 39, "ДАТА_РАЗМЕЩЕНИЯ_ОБЪЯВЛЕНИЯ")
	       self.ws.write(0, 40, "ЗАГОЛОВОК")
	       self.ws.write(0, 41, "ДАТА_ПАРСИНГА")
	       self.ws.write(0, 42, "МЕСТОПОЛОЖЕНИЕ")
	      	       
	       
	       self.result= 1
	       
                
                
                
                
	  def task_generator(self):
	       for x in range(250):
                    yield Task ('post',url=self.f+'?page=%d'%x,refresh_cache=True,network_try_count=100)
               yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@class="offer__details-link"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,network_try_count=100)
	       #yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	      
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//li[@class="new-pager__navigation-next"]/a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except DataNotFound:
		    print('*'*100)
		    print '!!!','NO PAGE NEXT','!!!'
		    print('*'*100)            
	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//div[@class="offer-detail__district"]/a[contains(text(),"район")]').text()
	       except IndexError:
	            mesto =''
		    
	       try:
	            punkt = grab.doc.select(u'//div[@class="offer-detail__city"]/a').text()
	       except IndexError:
	            punkt = ''	       
		
               try:
                    ter =  grab.doc.select(u'//div[@class="offer-detail__sublocality"]/a[contains(text(),"район")]').text()
               except IndexError:
                    ter =''
               try:
		    try:
                         uliza = grab.doc.select(u'//div[@class="offer-detail__address"]/a').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//div[@class="offer-detail__address"]').text()
               except IndexError:
                    uliza = ''
               try:
		    dom = 'Квартира'
               except IndexError:
                    dom = ''
		    
               try:
                    tip = grab.doc.select(u'//div[@class="offer-detail__price-per-square-rur"]').text()
               except IndexError:
                    tip = ''
               try:
                    naz = re.sub('[^\d]','',grab.doc.select(u'//h1').text().split(', ')[0])
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//div[contains(text(),"Жилая площадь")]/following-sibling::div').text()
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//div[@class="offer-detail__price-rur"]').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//div[contains(text(),"Общая площадь")]/following-sibling::div').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//div[contains(text(),"Этаж")]/following-sibling::div[contains(text(),"из")]').text().split(u' из ')[0]
               except IndexError:
                    ohrana =''
               try:
		    try:
                         gaz =  grab.doc.select(u'//div[contains(text(),"Этажность")]/following-sibling::div').text()
		    except IndexError:
			 gaz =  grab.doc.select(u'//div[contains(text(),"Этаж")]/following-sibling::div[contains(text(),"из")]').text().split(u' из ')[1]
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//div[contains(text(),"Год постройки")]/following-sibling::div').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//a[@class="offer-detail__metro-link"]').text()
               except IndexError:
                    kanal =''
               try:
                    elek = grab.doc.select(u'//div[contains(text(),"Удаленность")]/following-sibling::div').text()
               except IndexError:
                    elek =''
               try:
                    teplo = grab.doc.select(u'//h1').text()
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//div[@class="offer-detail__section-item section_type_additional-info"]').text() 
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//div[contains(text(),"частное лицо")]/preceding-sibling::div[@class="offer-detail__contact-name"]').text()
               except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//div[contains(text(),"агентство")]/preceding-sibling::div[@class="offer-detail__contact-name"]').text()
               except IndexError:
                    comp = ''
               
	       try: 
	            con = [ (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
		              (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]
		              
		    t= grab.doc.select(u'//div[@class="offer-detail__refresh"]').text()
		    if t.find(u'назад')>=0:
			 dt = u'вчера'
		    else:
		         dt= t
		    data = reduce(lambda dt, r: dt.replace(r[0], r[1]), con, dt)
	       except IndexError:
		    data=''
	       
	       try:
                    phone = re.sub('[^\d\+]','',grab.doc.select(u'//div[@class="offer-detail__contact-phone-wrapper"]').attr('data-phone'))
               except IndexError:
	            phone = ''
		    
		    
	       try:
		    ohr = grab.doc.select(u'//div[contains(text(),"Площадь кухни")]/following-sibling::div').text()
	       except IndexError:
		    ohr =''
		    
		    
	       if grab.doc.select(u'//div[contains(text(),"кондиционирования")]').exists() == True:
		    cond = 'Есть'
	       else:
	            cond = ''   
	     
	       try:
		    inet = grab.doc.select(u'//div[contains(text(),"Материал дома")]/following-sibling::div').text()
	       except IndexError:
		    inet =''
	       try:
		    lat = grab.doc.select(u'//div[contains(text(),"Расположение комнат")]/following-sibling::div').text()
	       except IndexError:
		    lat =''
	  
	       try:
		    lng = grab.doc.select(u'//div[contains(text(),"Высота потолков")]/following-sibling::div').text()
	       except IndexError:
		    lng =''		    
	       try:
		    lini = grab.doc.select(u'//div[contains(text(),"Окна")]/following-sibling::div').text()#.split(u' из ')[1]
	       except IndexError:
		    lini =''
	  
	  
	       try:
		    usl = grab.doc.select(u'//div[contains(text(),"Общепит в здании")]/following-sibling::div').text() 
	       except IndexError:
		    usl = ''
		    
	       try:
		    otd = grab.doc.select(u'//div[contains(text(),"Состояние")]/following-sibling::div').text() 
	       except IndexError:
	            otd = ''		    
	  
	       try:
		    park = grab.doc.select(u'//span[contains(text(),"Парковка")]/following::span[1]').text() 
	       except IndexError:
		    park = ''	       
		    
	       try:
		    lin = []
		    for em in grab.doc.select(u'//div[@class="offer-detail__location-block"]/div'):
	                 urr = em.text()
	                 lin.append(urr) 
	            rasp = ",".join(lin).replace(u'на карте,','')
	       except IndexError:
	            rasp =''	       
          
	       
	       try:
		    if 'prodazha' in task.url:
		         oper = u'Продажа'
		    elif 'arenda' in task.url:
		         oper = u'Аренда'
	       except IndexError:
		    oper = ''	       
     
	
               projects = {'sub': self.sub,
	                  'adress': mesto,
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
	                   'oper':oper,
	                   'lico':lico,
	                   'company': comp,
	                   'ohra':ohr,
	                   'condi':cond,
	                   'internet':inet,
	                   'shir': lat,
	                   'dol': lng,
	                   'linii': lini,
	                   'mesto': rasp.replace(kanal,''),
	                   'uslugi':usl,
	                   'sos': otd,
	                   'parkov': park,
	                   'data':data.replace(u'только что',datetime.today().strftime('%d.%m.%Y'))}
	                   
	  
	  
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
	       print  task.project['klass']
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
	       print  task.project['ohra']
	       print  task.project['condi']
	       print  task.project['internet']
	       print  task.project['shir']
	       print  task.project['dol']
	       print  task.project['linii']
	       print  task.project['uslugi']
	       print  task.project['sos']
	       print  task.project['parkov']
	       print  task.project['mesto']
	      
	  
	       
	       
     
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['adress'])
	       self.ws.write(self.result, 3, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 10, task.project['dom'])
	       self.ws.write(self.result, 13, task.project['tip'])
	       self.ws.write(self.result, 14, task.project['naz'])
	       self.ws.write(self.result, 16, task.project['klass'])
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 15, task.project['plosh'])
	       self.ws.write(self.result, 19, task.project['ohrana'])
	       self.ws.write(self.result, 20, task.project['gaz'])
	       self.ws.write(self.result, 22, task.project['voda'])
	       self.ws.write(self.result, 7, task.project['kanaliz'])
	       self.ws.write(self.result, 8, task.project['electr'])
	       self.ws.write(self.result, 40, task.project['teplo'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, u'RealtyMag.RU')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write_string(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 41, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 11, task.project['oper'])
	       self.ws.write(self.result, 17, task.project['ohra'])
	       #self.ws.write(self.result, 39, task.project['condi'])
	       self.ws.write(self.result, 21, task.project['internet'])
	       self.ws.write(self.result, 23, task.project['shir'])
	       self.ws.write(self.result, 29, task.project['dol'])
	       self.ws.write(self.result, 27, task.project['linii'])
	       #self.ws.write(self.result, 42, task.project['uslugi'])
	       self.ws.write(self.result, 28, task.project['sos'])
	       #self.ws.write(self.result, 37, task.project['parkov'])
	       self.ws.write(self.result, 42, task.project['mesto'])
	       
	       
	       
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)#+'/'+str(self.num)
	       print 'Tasks - %s' % self.task_queue.size() 
	       print '***',i+1,'/',len(l),'***'
	       print task.project['oper'] 
	       print('*'*100)
	       self.result+= 1
	       
	       
	       
	       #if str(self.result) == str(self.num):
		    #self.stop()	       
	 

     bot = Mag_Com(thread_number=5, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=500, connect_timeout=500)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     #command = 'mount -a'
     #p = os.system('echo %s|sudo -S %s' % ('1122', command))
     #print p
     #time.sleep(2)
     bot.workbook.close()
     print('Done!')
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
       
     
     
     
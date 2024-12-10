#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError,GrabTooManyRedirectsError 
import logging
import base64
import time
import os
import math
import json
import re
from datetime import datetime
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



#g.proxylist.load_file(path='/home/oleg/Proxy/tipa.txt',proxy_type='http')

i = 0

l= open('Links/Zem.txt').read().splitlines()
dc = len(l)
page = l[i]

while True:
     print '********************************************',i+1,'/',dc,'*******************************************'     
     class IRR_Zem(Spider):
          def prepare(self):
	       #self.count = 1 
	       self.f = page
	       #self.link =l[i]
	       #for p in range(1,150):
		    #try:
			 #time.sleep(1)
			 #g = Grab(timeout=10, connect_timeout=20)
			 #g.proxylist.load_file(path='../ivan.txt',proxy_type='http')
			 #g.go(self.f)
			 #city = [ (u'кой',u'кая'),(u'области',u'область'),(u'ком',u'кий'),(u'Москве',u'Москва'),(u'Петербурге',u'Петербург'),(u'крае',u'край'),(u'республике ','')]
                         ##dt = g.doc.select(u'//span[@itemprop="name"]').text().replace('Все объявления в ','').replace('Все объявления во ','').replace('/','-')
			 #dt = g.doc.select(u'//title').text().split(' | ')[0].replace('Земельные участки в ','').replace('/','-')
                         #self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), city, dt)
                         ##self.num= re.sub('[^\d]', '',g.doc.rex_text(u'var adverts_count = (.*?);'))
			 ##try:
			      ##self.num= re.sub('[^\d]', '',g.doc.select(u'//div[@class="listingStats"]').text().split('из ')[1])
			 ##except IndexError:
			      ##self.num = '1'
			 ##self.pag = int(math.ceil(float(int(self.num))/float(30)))
			 #print self.sub
			 #del g
			 #break
		    #except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError,GrabTooManyRedirectsError):
			 #print g.config['proxy'],'Change proxy'
			 #g.change_proxy()
			 #del g
			 #continue
               #else:
		    #self.sub = ''
		    
	       self.workbook = xlsxwriter.Workbook(u'Zem/IRR_'+str(i+1)+u'_Земля_Продажа.xlsx')
               self.ws = self.workbook.add_worksheet(u'Irr_ЗЕМЛЯ')
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
	       self.ws.write(0, 28, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 29, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 30, u"ВИД_ПРАВА")
	       self.ws.write(0, 31, u"ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
	       self.conv = [(u' августа',u'.08.2019'), (u' июля',u'.07.2019'),
			 (u' мая',u'.05.2019'),(u' июня',u'.06.2019'),
			 (u' марта',u'.03.2019'),(u' апреля',u'.04.2019'),
			 (u' января',u'.01.2019'),(u' декабря',u'.12.2018'),
			 (u' сентября',u'.09.2019'),(u' ноября',u'.11.2018'),
			 (u' февраля',u'.02.2018'),(u' октября',u'.10.2018'), 
			 (u'сегодня',datetime.today().strftime('%d.%m.%Y'))]	       
	       self.result= 1
	       
            
            
            
              
    
          def task_generator(self):
               #for x in range(1,self.pag+1):
                    #link = self.f+'page'+str(x)+'/'
                    #yield Task ('post',url=link.replace(u'page1/',''),refresh_cache=True,network_try_count=100)
	       yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)

	  def task_post(self,grab,task):
	       if grab.doc.select(u'//h4[contains(text(),"Предложения рядом")]').exists()==True:
		    links = grab.doc.select(u'//h4[contains(text(),"Предложения рядом")]/preceding::a[contains(@class,"listing")]')
               else:
                    links = grab.doc.select(u'//a[@class="listing__itemTitle"]')
	       for elem in links:
                    ur = grab.make_url_absolute(elem.attr('href'))
                    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
		    
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//li[contains(@class,"active")]/following-sibling::li[1]/a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print 'no_page'		    

	  def task_item(self, grab, task):
	       
	       try:
		    sub =  grab.doc.rex_text('"address_region":"(.*?)"').decode("unicode_escape").replace(u'russia\/moskva-region\/',u'Москва').replace(u'russia\/tatarstan-resp\/',u'Татарстан')	     
	       except (IndexError,TypeError,ValueError):
		    sub = ''
	       except KeyError:
		    sub = u'Санкт-Петербург'	       
	       try:
		    mesto = grab.doc.select(u'//i[@class="icon icon_spot"]/following-sibling::div[1]').text()
	       except DataNotFound:
	            mesto = ''	      
	      
	       try:
                    ra = grab.doc.select(u'//title').text()
                    t=0
                    for w in ra.split(', '):
	                 t+=1
	                 if w.find(u'р-н')>=0:
	                      ray = ra.split(', ')[t-1]
	                      break
			 elif w.find(u'район')>=0:
			      ray = ra.split(', ')[t-1]
			      break
                         else:
	                      ray =''
               except IndexError:
                    ray = ''
		    
               ter = ''
		    
               try:
		    uliza = grab.doc.select(u'//li[contains(text(),"Направление:")]').text().split(': ')[1]
               except IndexError:
                    uliza = ''

		  
	       try:
		    punkt = grab.doc.rex_text(u'address_city":"(.*?)"').decode("unicode_escape")
		    #print punkt
	       except (IndexError,TypeError,ValueError,KeyError):
		    punkt = ''
		    
              
		  
	       try:
		    trassa = grab.doc.select(u'//li[contains(text(),"Шоссе:")]').text().split(': ')[1]
	       except IndexError:
		    trassa = ''
		  
	       try:
		    udal = grab.doc.select(u'//li[contains(text(),"Удаленность:")]').text().split(': ')[1]
	       except IndexError:
		    udal = ''
		  
		  
	       try:
		    price = grab.doc.select(u'//div[@class="productPage__price js-contentPrice"]').text()
		#print price + u' руб'	    
	       except IndexError:
		    price = ''
		  
		  
	       try:
		    plosh = grab.doc.select(u'//li[contains(text(),"Площадь участка:")]').text().split(': ')[1]
		 #print rayon
	       except IndexError:
		    plosh = ''
		  
	       try:
                    categoria = grab.doc.select(u'//li[contains(text(),"Категория земли:")]').text().split(': ')[1]
	       except IndexError:
		    categoria = ''
		    
	       try:
	            vid = grab.doc.select(u'//li[contains(text(),"Вид разрешенного использования:")]').text().split(': ')[1]
	       except IndexError:
	            vid = ''		    
		  
	      
		  
	       try:
		    gaz = grab.doc.select(u'//li[contains(text(),"Газ")]').text().replace(u'Газ',u'есть')
	       except IndexError:
		    gaz =''
		  
	        
	       try:
		    elekt = grab.doc.select(u'//li[contains(text(),"Электричество")]').text().replace(u'Электричество (подведено)',u'есть')
	       except IndexError:
		    elekt =''
                 
	       try:
                    ohrana = grab.doc.select(u'//li[contains(text(),"Охрана")]').text().replace(u'Охрана',u'есть')
               except IndexError:
		    ohrana =''
		  
	       
		  
	       try:
		    opis = grab.doc.select(u'//h4[contains(text(),"Описание")]/following-sibling::p').text() 
	       except IndexError:
		    opis = ''
		  
	       try:
                    try:
	                 lico = grab.doc.select(u'//input[@name="contactFace"]').attr('value')
                    except IndexError:
	                 lico = grab.doc.select(u'//div[@class="productPage__infoTextBold productPage__infoTextBold_inline"]').text()
               except IndexError:
	            lico = ''
		  
	       try:
		    com = grab.doc.select(u'//div[@class="productPage__infoTextBold productPage__infoTextBold_inline"]/a[contains(@href,"russia")]').text()
		#print rayon
	       except IndexError:
		    com = ''
		  
		  
	       try:
		    d = grab.doc.select(u'//div[@class="productPage__createDate"]').text().split(', ')[0]
		    data = reduce(lambda d, r: d.replace(r[0], r[1]), self.conv, d)[:10]
		#print rayon
	       except IndexError:
		    data = ''
	       try:
		    data1 = datetime.strptime(grab.doc.rex_text(u'date_create":"(.*?)"}').split(' ')[0].replace('-','.'), '%Y.%m.%d')
		    #print rayon
	       except IndexError:
	            data1 = ''		    
		  
	       try:
                    prava = grab.doc.select(u'//li[contains(text(),"Вид права:")]').text().split(': ')[1]
		#print rayon
	       except IndexError:
		    prava = ''
		  
		  
	       try:
		    phone = re.sub('[^\d]', '',base64.b64decode(grab.doc.select('//input[@name="phoneBase64"]').attr('value')))
	       except (AttributeError,IndexError):
		    phone = ''  
		  
	       
		  
	      
	      
	      
	      
	      
	       projects = {'sub': sub,
		            'rayon': ray,
		            'punkt': punkt,
	                    'teritor': ter,
	                    'ulica': uliza,
		            'phone': phone,
	                     'vid': vid,
		            'price': price,
		            'opis': opis,
		            'url': task.url,
		            'trassa': trassa,
		            'udal': udal,
		            'ploshad': plosh,
		            'categoria': categoria,
		            'gaz': gaz,
		            'elekt': elekt,
		            'ohrana': ohrana,
		            'lico':lico.replace(com,''),
	                    'mesto':mesto,
		            'com':com,
		            'dataraz': data,
	                    'dataraz1': data1.strftime('%d.%m.%Y'),
		           'prava':prava
	                   
		         }
	
	
	
	       yield Task('write',project=projects,grab=grab)
	
  
	
	
	
	
	  def task_write(self,grab,task):
	      
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['opis']
	       print  task.project['price']
	       print  task.project['trassa']
	       print  task.project['udal']
	       print  task.project['ploshad']
	       print  task.project['categoria']
	       print  task.project['vid']
	       print  task.project['gaz']
	       print  task.project['elekt']
	       print  task.project['ohrana']
	       print  task.project['lico']
	       print  task.project['com']
	       print task.project['url']
	       print  task.project['phone']	       
	       print  task.project['dataraz']
	       print  task.project['dataraz1']
	       print  task.project['mesto']
	       print  task.project['prava']
	       
	       
	
	
	
	
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 6, task.project['ulica'])
	       self.ws.write(self.result, 7, task.project['trassa'])
	       self.ws.write(self.result, 8, task.project['udal'])
	       self.ws.write(self.result, 10, task.project['price'])
	       self.ws.write(self.result, 12, task.project['ploshad'])
	       self.ws.write(self.result, 13, task.project['categoria'])
	       self.ws.write(self.result, 14, task.project['vid'])
	       self.ws.write(self.result, 15, task.project['gaz'])
	       self.ws.write(self.result, 18, task.project['elekt'])
	       self.ws.write(self.result, 20, task.project['ohrana'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['com'])
	       self.ws.write(self.result, 28, task.project['dataraz'])
	       self.ws.write(self.result, 30, task.project['prava'])
	       self.ws.write(self.result, 23, u'Из рук в руки')
	       self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 9, u'Продажа')
	       self.ws.write(self.result, 31, task.project['dataraz1'])
	       self.ws.write(self.result, 32, task.project['mesto'])
	       print('*'*50)	
	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print '*',i+1,'/',dc,'*'
               print('*'*50)
	       self.result+= 1

     
     bot = IRR_Zem(thread_number=5,network_try_limit=1000)
     #bot.setup_queue(backend='mongo', database='irrzem',host='192.168.10.200')
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
	  break
time.sleep(5)
os.system("/home/oleg/pars/irr/com.py")

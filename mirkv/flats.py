#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
import math
import random
#import json
import time
import os
from grab import Grab
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


i = 0
l= open('links/Kv_all.txt').read().decode('cp1251').splitlines()
dc = len(l)
page = l[i]

     
while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class MK_Kv(Spider):
	  def prepare(self):
	       self.f = page
	       for p in range(1,21):
		    try:
			 time.sleep(1)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 self.sub = self.f.split('/')[3].replace('+',' ')
			 if u'Комнаты' in self.f:
			      self.tip_ob = u'Комнатa'
			 else:
			      self.tip_ob = u'Квартира'
			 print self.sub,self.tip_ob
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError,ValueError):
			 print g.config['proxy'],'Change proxy'
			 print str(p)
			 del g
			 continue
	       else:
		    self.pag = 0
		    self.num=0
		    self.stop()	
		    
	       self.workbook = xlsxwriter.Workbook(u'flat/Mirkvartir_Жилье_'+str(i)+'.xlsx')
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
	       self.ws.write(0, 22, u"СРОК_СДАЧИ")
	       self.ws.write(0, 23, u"РАСПОЛОЖЕНИЕ_КОМНАТ")
	       self.ws.write(0, 24, u"БАЛКОН")
	       self.ws.write(0, 25, u"ЛОДЖИЯ")
	       self.ws.write(0, 26, u"САНУЗЕЛ")
	       self.ws.write(0, 27, u"ОКНА")
	       self.ws.write(0, 28, u"СОСТОЯНИЕ")
	       self.ws.write(0, 29, u"ВЫСОТА_ПОТОЛКОВ")
	       self.ws.write(0, 30, u"ЛИФТ")
	       self.ws.write(0, 31, u"РЫНОК")
	       self.ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
	       self.ws.write(0, 33, u"ОПИСАНИЕ")
	       self.ws.write(0, 34, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 35, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 36, u"ТЕЛЕФОН")
	       self.ws.write(0, 37, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 38, u"КОМПАНИЯ")
	       self.ws.write(0, 39, u"ДАТА_РАЗМЕЩЕНИЯ_ОБЪЯВЛЕНИЯ")
	       self.ws.write(0, 40, u"ДАТА_ПАРСИНГА")
	       #self.ws.write(0, 41, u"ДОП._ИНФОРМАЦИЯ")
	       
	       self.result= 1
	    
	    
	    
	      
     
	  def task_generator(self):
	       for x in range(100):
                    yield Task ('post',url=self.f+'?p=%d'%x,refresh_cache=True,network_try_count=100)
	       yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
	
	    
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//a[@class="offer-title"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	  
   
	  def task_item(self, grab, task):
	       
	       try:
		    ray = grab.doc.select(u'//p[@class="address"]/a[contains(text(),"р-н")]').text()
	       except IndexError:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[3]
	       except IndexError:
		    punkt = ''
	       try:
		    ter= grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[1].replace(u'цена','')
	       except IndexError:
		    ter =''
	       try:
		    uliza = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[4].replace(u'цена','')
	       except IndexError:
		    uliza = ''
	       try:
		    dom = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[5].replace(u'цена','')
	       except IndexError:
		    dom = ''
		     
		 
	       try:
		    metro = grab.doc.rex_text(u'subwayName(.*?),')[3:][:-1]
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//p[@class="address"]/following-sibling::p/small').text()
		 #print rayon
	       except IndexError:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//div[@class="complex-info"]/h3/a[1]').text()
	       except IndexError:
		    metro_tr = ''
		    
		   
	       try:
		    price = grab.doc.select(u'//div[@class="price m-all"]').text()
		 #print price + u' руб'	    
	       except IndexError:
		    price = ''
		   
	       try:
		    price_m = grab.doc.select(u'//div[@class="price m-m2"]').text()#.split(u'.')[0]
	       except IndexError:
		    price_m = ''
		     
	       try:
		    kol_komnat = re.findall(u'"Комнаты","text":"(.*?)"}',grab.doc.rex_text(u'infoViewModel(.*?)secondItemsColumn').split('firstItemsColumn')[1])[0].split('-')[0]
	       except IndexError:
		    kol_komnat = ''
    
	       try:
		    plosh_ob = re.findall(u'"Площадь","text":"(.*?)"}',grab.doc.rex_text(u'infoViewModel(.*?)secondItemsColumn').split('firstItemsColumn')[1])[0].split(', ')[0]
		  #print rayon
	       except IndexError:
		    plosh_ob = ''
     
	       try:
		    plosh_gil =  re.findall(u'"Площадь","text":"(.*?)"}',grab.doc.rex_text(u'infoViewModel(.*?)secondItemsColumn').split('firstItemsColumn')[1])[0].split(', ')[2]
	       except IndexError:
	            plosh_gil = ''
		     
	       try:
		    plosh_kuh = re.findall(u'"Площадь","text":"(.*?)"}',grab.doc.rex_text(u'infoViewModel(.*?)secondItemsColumn').split('firstItemsColumn')[1])[0].split(', ')[1]
	       except IndexError:
	            plosh_kuh = ''
		  
	       try:
		    plosh_com = grab.doc.select(u'//label[contains(text(),"Комнаты:")]/following-sibling::p/br/following-sibling::text()').text()
	       except IndexError:
		    plosh_com = ''
		    
	       try:
		    et = re.findall(u'"Этаж","text":"(.*?)"}',grab.doc.rex_text(u'infoViewModel(.*?)secondItemsColumn').split('firstItemsColumn')[1])[0].split(' из ')[0]
	       except IndexError:
		    et = '' 
		   
	       try:
                    etagn = re.findall(u'"Этаж","text":"(.*?)"}',grab.doc.rex_text(u'infoViewModel(.*?)secondItemsColumn').split('firstItemsColumn')[1])[0].split(' из ')[1]
	       except IndexError:
		    etagn = ''
		     
	       try:
		    mat = grab.doc.select(u'//span[contains(text(),"Материал стен")]/following-sibling::strong').text()
		 #print rayon
	       except IndexError:
		    mat = '' 
		   
	       try:
		    god = grab.doc.select(u'//span[contains(text(),"Срок сдачи")]/preceding-sibling::strong').text()
	       except IndexError:
		    god = ''
		     
	       try:
		    balkon = grab.doc.select(u'//th[contains(text(),"Балкон:")]/following-sibling::td[contains(text(),"балк")]').text()#.replace(u'нет','')
		 #print rayon
	       except IndexError:
		    balkon = ''
		   
	       try:
		    lodg = grab.doc.select(u'//span[contains(text(),"Планировка")]/following-sibling::strong').text().split(u', ')[0]
		 #print rayon
	       except IndexError:
		    lodg = ''
		   
	       try:
		    sanuzel = grab.doc.select(u'//span[contains(text(),"Планировка")]/following-sibling::strong').text().split(u', ')[1]
	       except IndexError:
		    sanuzel = ''
		     
		     
	       try:
		    okna = grab.doc.select(u'//span[contains(text(),"Отделка")]/following-sibling::strong').text()
	       except IndexError:
		    okna = ''
		   
	       try:
		    lift = grab.doc.select(u'//span[contains(text(),"Высота потолков")]/following-sibling::strong').text()
	       except IndexError:
		    lift = ''
		  
	       try:
		    rinok = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[2]
		    if u'новостройке' in rinok:
			 rinok = u'Новостройка'
		    else:
			 rinok = u'Вторичка'
	       except IndexError:
		    rinok = ''
		   
	       try:
		    kons = grab.doc.select(u'//p[@class="address"]').text()
	       except IndexError:
		    kons = ''
		     
	       try:
		    opis = grab.doc.select(u'//div[@class="l-object-description"]/p').text() 
	       except IndexError:
		    opis = ''
     
	       try:
		    lico = grab.doc.select(u'//div[@class="seller-info"]/p/strong').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//span[contains(text(),"Агентство недвижимости")]/preceding-sibling::strong').text()
		 #print rayon
	       except IndexError:
		    comp = ''
		    
	       try:
                    try:
	                 data = grab.doc.select(u'//div[@class="l-object-dates"]/p[2]').text().split(u' размещено ')[1].split(u' в ')[0].replace(u'сегодня', (datetime.today().strftime('%d.%m.%Y')))
                    except IndexError:
	                 data = grab.doc.select(u'//div[@class="dates"]').text().split(u' размещено ')[1].split(u' в ')[0].replace(u'сегодня', (datetime.today().strftime('%d.%m.%Y')))
               except IndexError:
	            data = ''
		    
     
	       projects = {'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
		           'metro': metro,
		           'udall': metro_min,
		           'tran': metro_tr,
		           'object': self.tip_ob,
		           'cena': price,
		           'cena_m': price_m,
		           'col_komnat': kol_komnat,
		           'plosh_ob':plosh_ob,
		           'plosh_gil': plosh_gil,
		           'plosh_kuh': plosh_kuh,
		           'plosh_com': plosh_com,
		           'etach': et,
		           'etashost': etagn,
		           'material': mat,
		           'god_postr': god,
	                   'phone': random.choice(list(open('../phone.txt').read().splitlines())),
		           'balkon': balkon,
		           'logia': lodg,
		           'uzel':sanuzel,
		           'okna': okna,
		           'lift':lift,
		           'rinok': rinok,
		           'kons':kons,
		           'opis':opis,
		           'url':task.url,
		           'lico':lico.replace(comp,''),
		           'company':comp,
		           'data':data.replace('.18','.2018').replace('.19','.2019')}
	     
	       yield Task('write',project=projects,grab=grab)
	       
	  def task_write(self,grab,task):
	       
	       print('*'*50)	       
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['metro']
	       print  task.project['udall']
	       print  task.project['tran']
	       print  task.project['cena']
	       print  task.project['cena_m']
	       print  task.project['col_komnat']
	       print  task.project['plosh_ob']
	       print  task.project['plosh_gil']
	       print  task.project['plosh_kuh']
	       print  task.project['plosh_com']
	       print  task.project['etach']
	       print  task.project['etashost']
	       print  task.project['material']
	       print  task.project['god_postr']
	       print  task.project['balkon']
	       print  task.project['logia']
	       print  task.project['uzel']
	       print  task.project['okna']
	       print  task.project['lift']
	       print  task.project['rinok']
	       print  task.project['kons']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       #print  task.project['tip_prod']
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 11,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 7,task.project['metro'])
	       self.ws.write(self.result, 8,task.project['udall'])
	       self.ws.write(self.result, 6,task.project['tran'])
	       self.ws.write(self.result, 10,task.project['object'])
	       #self.ws.write(self.result, 11,oper)
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 13, task.project['cena_m'])
	       self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 15, task.project['plosh_ob'])
	       self.ws.write(self.result, 16, task.project['plosh_gil'])
	       self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       self.ws.write(self.result, 18, task.project['plosh_com'])
	       self.ws.write(self.result, 19, task.project['etach'])
	       self.ws.write(self.result, 20, task.project['etashost'])
	       self.ws.write(self.result, 21, task.project['material'])
	       self.ws.write(self.result, 22, task.project['god_postr'])
	       self.ws.write(self.result, 24, task.project['balkon'])
	       self.ws.write(self.result, 25, task.project['logia'])
	       self.ws.write(self.result, 26, task.project['uzel'])
	       self.ws.write(self.result, 28, task.project['okna'])
	       self.ws.write(self.result, 29, task.project['lift'])
	       self.ws.write(self.result, 31, task.project['rinok'])
	       self.ws.write(self.result, 32, task.project['kons'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, u'MIRKVARTIR.RU')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 40, datetime.today().strftime('%d.%m.%Y'))
	       #self.ws.write(self.result, 41, task.project['tip_prod'])
	       
	       print('*'*50)
	       print self.sub
	      
	       print 'Ready - '+str(self.result)#+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '**',i+1,'/',dc,'***'
	       print  task.project['teritor']
	       print  task.project['object']
	       print('*'*50)	       
	       self.result+= 1
	       
	   
	       #if str(self.result) == str(self.num):
		    #self.stop()
	
	
	       
     bot = MK_Kv(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=500, connect_timeout=500)
     try:
	  bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     time.sleep(1)
     bot.workbook.close()
     print('Done')
     i=i+1
     try:
	  page = l[i]
     except IndexError:
	  break     
     
     
     
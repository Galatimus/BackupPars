#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import time
import math
import random
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
l= open('Links/urls.txt').read().splitlines()
dc = len(l)
page = l[i] 

oper = u'Продажа'




while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class Mlsn_Kv(Spider):
	  def prepare(self):
	       self.f = page
	       for p in range(1,16):
		    try:
			 time.sleep(2)
			 g = Grab(timeout=25, connect_timeout=65)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f+'/kvartiry/')
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
		    self.sub = ''
		    self.num = 1
		    self.pag = 1
		    self.stop()   
		    
	       
	       self.workbook = xlsxwriter.Workbook(u'flat/Move_Жилье_'+oper+'_'+str(i+1) + '.xlsx')
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
	       self.ws.write(0, 39, u"ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 40, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 41, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 42, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):
	       for x in range(1,self.pag+1):
		    yield Task ('post',url=self.f+'/kvartiry/?page=%d'%x,refresh_cache=True,network_try_count=100)

		
            
	  def task_post(self,grab,task):

	       for elem in grab.doc.select(u'//a[@class="search-item__title-link search-item__item-link"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
     
	  def task_item(self, grab, task):
	       
	       try:
		    sub = grab.doc.select(u'//li[@class="top-menu__item"]/span').text()
	       except IndexError:
		    sub = ''	       
	       try:
		    ray = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title, " р-н")]').text()
	       except IndexError:
		    ray = ''          
	       try:
		    try:
			 try:
			      try:
				   try:
					try:
					     try:
						  try:
						       punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"г.")]').text()
						  except IndexError:
						       punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"поселок")]').text()
					     except IndexError:
						  punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"д. ")]').text()
					except IndexError:
					     punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"сл.")]').text()
				   except IndexError:
					punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"пгт")]').text()
			      except IndexError:
				   punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"с.")]').text()
			 except IndexError:
			      punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"рп")]').text()
		    except IndexError:
			 punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"п.")]').text()
	       except IndexError:
		    punkt =''
		    
		    
	       try:
	            ter= grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "rayon")]').text()
	       except IndexError:
		    ter =''
	       try:
		    try:
			 try:
			      try:
				   try:
					uliza = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "ulica")]').text()
				   except IndexError:
					uliza = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "prospekt")]').text()
			      except IndexError:
				   uliza = grab.doc.select(u'//div[@class="geo-block__geo-info_second-line"]/span[1]').text()
			 except IndexError:
			      uliza = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "proezd")]').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "shosse")]').text()
	       except IndexError:
		    uliza =''
	       try:
		    try:
			 dom = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title, "д.")]').text()
		    except IndexError:
			 dom = grab.doc.select(u'//div[@class="geo-block__geo-info_second-line"]/span[2]').text()
	       except IndexError:
		    dom = ''
		     
	       try:
		    lin = []
		    for em in grab.doc.select(u'//li[@class="object-info__details-table_property"]/div[contains(@title, "г.")]'):
			 urr = em.text().replace(':','')
			 #print urr
			 lin.append(urr)
		    orentir = ",".join(lin)
	       except IndexError:
		    orentir = ''
		 
	       try:
		    metro = grab.doc.select(u'//ul[@class="geo-block__block-distance"]/li/a[contains(@href,"metro")]').attr('title')
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//li[@class="geo-block__block-distance_property geo-block__block-distance_walk-time"]').text()
		 #print rayon
	       except IndexError:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//div[contains(text(),"Тип комнат:")]/following-sibling::div').text()
	       except IndexError:
		    metro_tr = ''

	       try:
		    price = grab.doc.select(u'//div[contains(text(),"Цена:")]/following-sibling::div').text()
	       except IndexError:
		    price = ''
		   
	       try:
		    price_m = grab.doc.select(u'//div[contains(text()," цена за м²:")]/following-sibling::div').text()
	       except IndexError:
		    price_m = ''
		     
	       try:
		    kol_komnat = grab.doc.select(u'//div[contains(text(),"Количество комнат:")]/following-sibling::div').text()
		#print rayon
	       except IndexError:
		    kol_komnat = ''
     
               try:
		    plosh_ob = grab.doc.select(u'//div[contains(text(),"Общая площадь:")]/following-sibling::div').text()
	       except IndexError:
		    plosh_ob = ''
     
	       try:
		    plosh_gil = grab.doc.select(u'//div[contains(text(),"Жилая комната:")]/following-sibling::div').text()
		  #print rayon
	       except IndexError:
		    plosh_gil = ''
		     
	       try:
		    plosh_kuh = grab.doc.select(u'//div[contains(text(),"Площадь кухни:")]/following-sibling::div').text()
		  #print rayon
	       except IndexError:
		    plosh_kuh = ''
		  
	       try:
		    plosh_com = grab.doc.select(u'//span[contains(text(),"Тип квартиры")]/following-sibling::span').text()
	       except IndexError:
		    plosh_com = ''
		    
	       try:
		    et = grab.doc.select(u'//div[contains(text(),"Этаж:")]/following-sibling::div').text().split('/')[0]
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		    etagn = grab.doc.select(u'//div[contains(text(),"Этаж:")]/following-sibling::div').text().split('/')[1]
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''
		     
	       try:
		    mat = grab.doc.select(u'//div[contains(text(),"Тип здания:")]/following-sibling::div').text().replace(u'не указано','')
		 #print rayon
	       except IndexError:
		    mat = '' 
		   
	       try:
		    god = grab.doc.select(u'//div[contains(text(),"Год постройки:")]/following-sibling::div').text().replace(u'не указано','')
	       except IndexError:
		    god = ''
		     
	       try:
		    balkon = grab.doc.select(u'//div[contains(text(),"Тип балкона:")]/following-sibling::div').text()
		 #print rayon
	       except IndexError:
		    balkon = ''
		   
	       try:
		    lodg = grab.doc.select(u'//div[contains(text(),"Ремонт:")]/following-sibling::div').text().replace(u'не указано','')
		 #print rayon
	       except IndexError:
		    lodg = ''
		   
	       try:
		    sanuzel = grab.doc.select(u'//div[contains(text(),"Тип санузла:")]/following-sibling::div').text().replace(u'не указано','')
	       except IndexError:
		    sanuzel = ''
		     
		     
	       try:
		    okna = grab.doc.select(u'//div[contains(text(),"Вид из окна:")]/following-sibling::div').text().replace(u'не указано','')
	       except IndexError:
		    okna = ''
		   
	       try:
	            potolki = grab.doc.select(u'//div[contains(text(),"Высота потолков:")]/following-sibling::div').text()
	       except IndexError:
	            potolki =''
		   
	       try:
		    lift = grab.doc.select(u'//div[contains(text(),"Лифт:")]/following-sibling::div').text().replace(u'не указано','')
	       except IndexError:
		    lift = ''
		  
	       try:
		    rinok = grab.doc.select(u'//div[contains(text(),"Тип объявления:")]/following-sibling::div').text()
	       except IndexError:
		    rinok = ''
		   
	       try:
		    kons = re.sub(u'^.*(?=консьерж)','', grab.doc.select(u'//*[contains(text(), "консьерж")]').text())[:8].replace(u'консьерж',u'есть')
	       except IndexError:
		    kons = ''
		     
	       try:
		    opis = grab.doc.select(u'//h4[contains(text(),"Описание")]/following-sibling::div[1]').text() 
	       except IndexError:
		    opis = ''
		
	       try:
                    phone = grab.doc.select(u'//div[contains(text(),"Тип объекта:")]/following-sibling::div').text()
	       except IndexError:
		    phone = ''
		   
	       try:
		    try:
			 lico = grab.doc.select(u'//div[@class="block-user__name"]').text()
		    except IndexError:
			 lico = grab.doc.select(u'//a[@class="block-user__name"]').text()
	       except IndexError:
	            lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//div[@class="block-user__agency"]').text()
		 #print rayon
	       except IndexError:
		    comp = ''
		    
	       try:
		    data =  grab.doc.select(u'//div[@class="object-place__address"]').text()
	       except IndexError:
		    data = ''
	       projects = {'sub': sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
	                   'orentir': orentir,
		           'metro': metro,
		           'potolki': potolki,
		           'tran': metro_tr,
		           'cena': price,
	                   'dometro': metro_min,
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
		           'balkon': balkon,
		           'logia': lodg,
		           'uzel':sanuzel,
		           'okna': okna,
		           'lift':lift,
		           'rinok': rinok,
		           'kons':kons,
		           'opis':opis,
		           'url':task.url,
		           'phone':phone,
		           'lico':lico,
		           'company':comp,
		           'data':data }
	     
	       try:
		    link = task.url.replace('objects','objectsv3/printing')
		    yield Task('phone',url=link,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    pass
	       
	  def task_phone(self, grab, task):
	       try:
		    phone= re.sub('[^\d\,]','',grab.doc.select(u'//div[@class="phone"]').text())
	       except IndexError:
		    phone = random.choice(list(open('../phone.txt').read().splitlines()))
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
	       
	       print('*'*50)	       
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['orentir']
	       print  task.project['metro']
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
	       print  task.proj['phone']
	       print  task.project['kons']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['potolki']
	       print  task.proj['dataraz']
	       print  task.proj['dataob']
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 6,task.project['orentir'])
	       self.ws.write(self.result, 7,task.project['metro'])
	       self.ws.write(self.result, 29,task.project['potolki'])
	       self.ws.write(self.result, 23,task.project['tran'])
	       self.ws.write(self.result, 8,task.project['dometro'])
	       self.ws.write(self.result, 11,oper)
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 13, task.project['cena_m'])
	       self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 15, task.project['plosh_ob'])
	       self.ws.write(self.result, 16, task.project['plosh_gil'])
	       self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       self.ws.write(self.result, 23, task.project['plosh_com'])
	       self.ws.write(self.result, 19, task.project['etach'])
	       self.ws.write(self.result, 20, task.project['etashost'])
	       self.ws.write(self.result, 21, task.project['material'])
	       self.ws.write(self.result, 22, task.project['god_postr'])
	       self.ws.write(self.result, 24, task.project['balkon'])
	       self.ws.write(self.result, 28, task.project['logia'])
	       self.ws.write(self.result, 26, task.project['uzel'])
	       self.ws.write(self.result, 27, task.project['okna'])
	       self.ws.write(self.result, 30, task.project['lift'])
	       self.ws.write(self.result, 31, task.project['rinok'])
	       self.ws.write(self.result, 32, task.project['kons'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, u'MOVE.RU')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 10, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 42, task.project['data'])
	       self.ws.write(self.result, 41, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 36, task.proj['phone'])
	       self.ws.write(self.result, 39, task.proj['dataraz'])
	       self.ws.write(self.result, 40, task.proj['dataob'])	       
	       
	       print('*'*50)
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',dc,'***'
	       print oper
	       print  task.project['phone']
	       print('*'*50)
	           
	       self.result+= 1
	       
	       if self.result > 25000:
		    self.stop()

     
     bot = Mlsn_Kv(thread_number=10,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=1000, connect_timeout=1000)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
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
          break
time.sleep(5)
#os.system("/home/oleg/pars/move/zaga.py")
     
     

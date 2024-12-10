#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
import time
from datetime import datetime,timedelta
import xlsxwriter

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

workbook = xlsxwriter.Workbook(u'flats/Cian_Жилье.xlsx') 


class Cian_Kv(Spider):
     def prepare(self):
	  self.ws = workbook.add_worksheet()
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
	  self.ws.write(0, 30, u"СРОК_СДАЧИ")
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
	  l= open('cian_flats.txt').read().splitlines()
          self.dc = len(l)
          print self.dc
          for line in l:
	       yield Task ('item',url=line,network_try_count=100)
        
     
     def task_item(self, grab, task):
	  
	  try:
               sub = grab.doc.select(u'//address[contains(@class,"address")]').text().split(', ')[0]
          except IndexError:
               sub = ''
	  try:
	       try:
                    ray = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"р-н ")]').text()
	       except IndexError:
		    ray = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"округ")]').text()
          except IndexError:
               ray =''
	  try:
               if sub == u'Москва':
	            punkt= u'Москва'
	       elif sub == u'Санкт-Петербург':
	            punkt= u'Санкт-Петербург'
	       elif sub == u'Севастополь':
	            punkt= u'Севастополь'
	       else:
	            if  grab.doc.select(u'//address[contains(@class,"address")]/a[2][contains(text(),"р-н ")]').exists()==True:
                         punkt= grab.doc.select(u'//address[contains(@class,"address")]/a[3]').text()
                    elif grab.doc.select(u'//address[contains(@class,"address")]/a[3][contains(text(),"р-н ")]').exists()==True:
	                 punkt= grab.doc.select(u'//address[contains(@class,"address")]/a[2]').text()
                    else:
	                 punkt=grab.doc.select(u'//address[contains(@class,"address")]/a[2]').text()
          except IndexError:
               punkt = ''
		 
	  try:
               ter= grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"район")]').text()
          except IndexError:
               ter ='' 
	  try:
	       try:
		    try:
			 try:
			      try:
				   try:
					try:
					     uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"ул.")]').text()
					except IndexError:
					     uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"пер.")]').text()
				   except IndexError:
					uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"просп.")]').text()
			      except IndexError:
				   uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"ш.")]').text()
			 except IndexError:
			      uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"бул.")]').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"проезд")]').text()
               except IndexError:
			 uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"наб.")]').text()
          except IndexError:
               uliza = ''
	       
	  try:
	       dom = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(@href,"house")]').text()
	   #print rayon
	  except IndexError:
	       dom = ''
		
	    
	  try:
	       metro = grab.doc.select(u'//a[contains(@href,"metro")]/span').text()
	    #print rayon
	  except IndexError:
	       metro = ''
	      
	  try:
	       metro_min = grab.doc.select(u'//a[contains(@href,"metro")]/following-sibling::span').text().split('. ')[0].replace(', ','')
	    #print rayon
	  except IndexError:
	       metro_min = ''
	      
	  try:
	       metro_tr = grab.doc.select(u'//a[contains(@href,"metro")]/following-sibling::span').text().split('. ')[1]
          except IndexError:
	       metro_tr = ''
	       
	  try:
	       #if grab.doc.select(u'//div[@class="object_descr_title"]').text().find(u'комната') == 0:
                    #tip_ob = u'Комната'
	       #else:
               tip_ob = u'Квартира' 
	  except IndexError:
	       tip_ob = ''
	       
	  try:
	       if 'sale' in task.url:
                    oper = u'Продажа' 
               elif 'rent' in task.url:
	            oper = u'Аренда' 
	  except IndexError:
	       oper = ''
	      
	  try:
	       price = grab.doc.select(u'//span[@itemprop="price"]').text()
	    #print price + u' руб'	    
	  except IndexError:
	       price = ''
	      
	  try:
               price_m = grab.doc.select(u'//div[contains(@class,"price_per_meter")]').text()#.split(u'.')[0]
          except IndexError:
               price_m = ''
		
	  try:
	       kol_komnat = grab.doc.select(u'//h1').text().split('-')[0]
	   #print rayon
	  except IndexError:
	       kol_komnat = ''

	  

	  try:
	       plosh_ob = grab.doc.select(u'//div[contains(text(),"Общая")]/preceding-sibling::div').text()
	     #print rayon
	  except IndexError:
	       plosh_ob = ''

	  try:
	       plosh_gil = grab.doc.select(u'//div[contains(text(),"Жилая")]/preceding-sibling::div[1]').text()
	     #print rayon
	  except IndexError:
	       plosh_gil = ''
		
	  try:
	       plosh_kuh = grab.doc.select(u'//div[contains(text(),"Кухня")]/preceding-sibling::div[1]').text()
	     #print rayon
	  except DataNotFound:
	       plosh_kuh = ''
	     
	  try:
	       plosh_com = grab.doc.select(u'//span[contains(text(),"Площадь комнат")]/following-sibling::span[1]').text().replace(u'–','')
          except DataNotFound:
	       plosh_com = ''
	       
	  try:
	       et = grab.doc.select(u'//div[contains(text(),"Этаж")]/preceding-sibling::div[1]').text().split(u' из ')[0]
	    #print price + u' руб'	    
	  except IndexError:
	       et = '' 
	      
	  try:
	       etagn = grab.doc.select(u'//div[contains(text(),"Этаж")]/preceding-sibling::div[1]').text().split(u' из ')[1]
	    #print price + u' руб'	    
	  except IndexError:
	       etagn = ''
		
	  try:
	       mat = grab.doc.select(u'//span[contains(text(),"Тип дома")]/preceding-sibling::span[1]').text()
	    #print rayon
	  except IndexError:
	       mat = '' 
	      
	  try:
	       god = grab.doc.select(u'//div[contains(text(),"Построен")]/preceding-sibling::div').text()
	  except DataNotFound:
	       god = ''
		
	  try:
	       balkon = grab.doc.select(u'//li[contains(text(),"Балкон")]').text().replace(u'Балкон',u'есть')
	    #print rayon
	  except IndexError:
	       balkon = ''
	      
	  try:
	       lodg = grab.doc.select(u'//li[contains(text(),"Лоджия")]').text().replace(u'Лоджия',u'есть')
	    #print rayon
	  except IndexError:
	       lodg = ''
	      
	  try:
	       sanuzel = grab.doc.select(u'//span[contains(text(),"санузел")]/following-sibling::span').text()
	  except IndexError:
	       sanuzel = ''
		
		
	  try:
	       okna = grab.doc.select(u'//span[contains(text(),"Вид из окон")]/following-sibling::span[1]').text()
	  except IndexError:
	       okna = ''
	      
	  #try:
	    #potolki = grab.doc.select(u'//div[contains(text(),"Высота потолков:")]/following-sibling::div[@class="propertyValue"]').text()
	  #except DataNotFound:
	      #potolki = ''
	      
	  try:
	       lift = grab.doc.select(u'//div[contains(text(),"Срок сдачи")]/preceding-sibling::div').text()
	  except IndexError:
	       lift = ''
	     
	  try:
	       rinok = grab.doc.select(u'//span[contains(text(),"Тип жилья")]/following-sibling::span[1]').text()
	  except IndexError:
	       rinok = ''
	      
	  try:
	       kons = re.sub(u'^.*(?=консьерж)','', grab.doc.select(u'//*[contains(text(), "консьерж")]').text())[:8].replace(u'консьерж',u'есть')
	  except IndexError:
	       kons = ''
		
	  try:
	       opis = grab.doc.select(u'//p[@itemprop="description"]').text() 
	  except IndexError:
	       opis = ''
	   
	  try:
	       try:
                    phone = re.sub(u'[^\d\+]','',grab.doc.select(u'//a[@class="phone--1OSCA"]').text())
               except IndexError:
	            phone = re.sub(u'[^\d\+]','',grab.doc.rex_text(u'offerPhone(.*?),'))
	  except IndexError:
	       phone = ''
	      
	  try:
	       try:
                    lico = grab.doc.select(u'//div[contains(text(), "Собственник")]').text()
               except IndexError:
	            lico = grab.doc.select(u'//span[contains(text(), "Застройщик")]').text()
	  except IndexError:
	       lico = ''
	       
	  try:
	       try:
                    comp = grab.doc.select(u'//div[contains(text(), "Агентство недвижимости")]/preceding::h2').text()
               except IndexError:
	            comp = grab.doc.select(u'//h4[@class="realtor-card__subtitle"]').text()
	  except IndexError:
	       comp = ''
	       
	  try:
               data = re.sub(u'[^\d\-]','',grab.doc.rex_text(u'editDate(.*?)T')).replace('-','.')
          except IndexError:
               data = ''
	       
	  
	       
	  try:
	       tip_pr = grab.doc.select(u'//address[contains(@class,"address")]').text().replace(u'На карте','')
	    #print rayon
	  except IndexError:
	       tip_pr = ''
	      
	      
	  
	      
         
      
	  projects = {'sub': sub,
	              'rayon': ray,
	              'punkt': punkt,
	              'teritor': ter,
	              'ulica': uliza,
	              'dom': dom,
	              'metro': metro,
	              'udall': metro_min,
	              'tran': metro_tr,
	              'object': tip_ob,
	              'cena': price,
	              'cena_m': price_m,
	              'col_komnat': kol_komnat.split(', ')[0],
	              'plosh_ob':plosh_ob,
	              'plosh_gil': plosh_gil,
	              'plosh_kuh': plosh_kuh,
	              'plosh_com': plosh_com,
	              'etach': et,
	              'etashost': etagn,
	              'material': mat,
	              'balkon': balkon,
	              'logia': lodg,
	              'uzel':sanuzel,
	              'okna': okna,
	              'lift':lift,
	              'rinok': rinok,
	              'kons':kons,
	              'opis':opis,
	              'url':task.url,
	              'postr': god,
	              'phone':phone,
	              'lico':lico,
	              'company':comp,
	              'data':data,
	              'tip_prod':tip_pr,
	              'oper':oper
	              }
	
	
	
	  yield Task('write',project=projects,grab=grab)
	
	
	
	
	
	
     def task_write(self,grab,task):
	  if task.project['sub'] <> '': 
	       print('*'*50)	       
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['metro']
	       print  task.project['udall']
	       print  task.project['tran']
	       print  task.project['object']
	       print  task.project['cena']
	       print  task.project['cena_m']
	       print  task.project['col_komnat']
	       print  task.project['plosh_ob']
	       print  task.project['postr']
	       print  task.project['plosh_gil']
	       print  task.project['plosh_kuh']
	       print  task.project['plosh_com']
	       print  task.project['etach']
	       print  task.project['etashost']
	       print  task.project['material']
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
	       print  task.project['oper']
	       print  task.project['tip_prod']
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 7,task.project['metro'])
	       self.ws.write(self.result, 8,task.project['udall'])
	       self.ws.write(self.result, 9,task.project['tran'])
	       self.ws.write(self.result, 10,task.project['object'])
	       self.ws.write(self.result, 11,task.project['oper'])
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
	       self.ws.write(self.result, 22, task.project['postr'])
	       self.ws.write(self.result, 24, task.project['balkon'])
	       self.ws.write(self.result, 25, task.project['logia'])
	       self.ws.write(self.result, 26, task.project['uzel'])
	       self.ws.write(self.result, 27, task.project['okna'])
	       self.ws.write(self.result, 30, task.project['lift'])
	       self.ws.write(self.result, 31, task.project['rinok'])
	       self.ws.write(self.result, 32, task.project['kons'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, u'ЦИАН')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 40, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 41, task.project['tip_prod'])
	       
	       print('*'*50)
	       print 'Ready - '+str(self.result)+'/'+str(self.dc)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print('*'*50)
	       self.result+= 1
	       
	       
	       #if self.result > 50:
		    #self.stop()

	 
     
bot = Cian_Kv(thread_number=6,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
try:
     bot.run()
except KeyboardInterrupt:
     pass
workbook.close()
print('Done!')

     
     
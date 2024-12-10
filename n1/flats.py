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



i = 0
l= open('Links/kv_all.txt').read().splitlines()
dc = len(l)
page = l[i] 






while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class Mlsn_Kv(Spider):
	  def prepare(self):
	       self.f = page     
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
			      
			 if 'kvartiry' in self.f:
			      self.tip_ob = u'Квартира' 
		         elif 'komnaty' in self.f:
			      self.tip_ob = u'Комната' 
			 self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="breadcrumbs"]/ul/li/span[contains(text(),"объявлени")]').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(25)))
			 self.sub = g.doc.select(u'//span[@class="search-2gen-geo-filter-caption-link__text"]').text()
			 print self.sub,self.oper,self.tip_ob,self.pag,self.num
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
		    
	       
	       self.workbook = xlsxwriter.Workbook(u'flat/N1_Жилье_'+str(i+1) + '.xlsx')
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
		    yield Task ('post',url=page+'?page=%d'%x,network_try_count=100)	       
	       #yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
		
            
	  def task_post(self,grab,task):

	       for elem in grab.doc.select(u'//a[@class="link"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
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
		    ray = grab.doc.select(u'//li[@class="geo-tags__item"]/a[contains(@href,"rayon")]/span[1]').text()
	       except IndexError:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//title').text().split('N1.RU ')[1]
	       except IndexError:
		    punkt = ''
	       try:
	            ter= grab.doc.select(u'//li[@class="geo-tags__item"]/a[contains(@href,"district")]/span[1]').text()
	       except IndexError:
		    ter =''
	       try:
		    uliza = grab.doc.select(u'//li[@class="geo-tags__item"]/a[contains(@href,"street")]/span[1]').text()
	       except (IndexError,TypeError):
		    uliza = ''
	       try:
		    dom = grab.doc.rex_text(u'house_number":"(.+?)",')
	       except (IndexError,TypeError):
		    dom = ''
		     
	       try:
		    orentir = grab.doc.select(u'//a[@class="geo-tags__item"][contains(@href,"microdistrict")]').text()
	       except IndexError:
		    orentir = ''              
		 
	       try:
		    metro = grab.doc.select(u'//a[@class="breadcrumbs-list__link"][contains(@href,"metro")]').text()
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//span[@class="card-living-content-location-metro__text _time"]').number()
		 #print rayon
	       except IndexError:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//span[@class="card-living-content-location-metro__text _time"]').text().split(u'минут ')[1]
	       except IndexError:
		    metro_tr = ''


	       try:
		    price = grab.doc.select(u'//div[@class="price"]').text()
	       except IndexError:
		    price = ''
		   
	       try:
		    price_m = grab.doc.select(u'//li[@class="card-living-content-deal-params__item"][1]').text()
	       except IndexError:
		    price_m = ''
		     
	       try:
		    kol_komnat = re.sub('[^0-9]','',grab.doc.select(u'//h1').text().split(u', ')[0])
		#print rayon
	       except IndexError:
		    kol_komnat = ''
     
	       
     
	       try:
		    plosh_ob = grab.doc.select(u'//span[contains(text(),"Общая площадь")]/following-sibling::span').text()
		  #print rayon
	       except IndexError:
		    plosh_ob = ''
     
	       try:
		    plosh_gil = grab.doc.select(u'//span[contains(text(),"Жилая площадь")]/following-sibling::span').text()
		  #print rayon
	       except IndexError:
		    plosh_gil = ''
		     
	       try:
		    plosh_kuh = grab.doc.select(u'//span[contains(text(),"Кухня")]/following-sibling::span').text()
		  #print rayon
	       except IndexError:
		    plosh_kuh = ''
		  
	       try:
		    plosh_com = grab.doc.select(u'//span[contains(text(),"Тип квартиры")]/following-sibling::span').text()
	       except IndexError:
		    plosh_com = ''
		    
	       try:
		    et = grab.doc.select(u'//span[contains(text(),"Этаж")]/following-sibling::span').text().split(' из ')[0]
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		    etagn = grab.doc.select(u'//span[contains(text(),"Этаж")]/following-sibling::span').text().split(' из ')[1]
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''
		     
	       try:
		    mat = grab.doc.select(u'//span[contains(text(),"Материал дома")]/following-sibling::span').text().replace(u'не указано','')
		 #print rayon
	       except IndexError:
		    mat = '' 
		   
	       try:
		    god = grab.doc.select(u'//span[contains(text(),"Год постройки")]/following-sibling::span').text().replace(u'не указано','')
	       except IndexError:
		    god = ''
		     
	       try:
		    balkon = grab.doc.select(u'//span[contains(text(),"Количество балконов")]/following-sibling::span').text().replace(u'нет','')
		 #print rayon
	       except IndexError:
		    balkon = ''
		   
	       try:
		    lodg = grab.doc.select(u'//span[contains(text(),"Состояние")]/following-sibling::span').text().replace(u'не указано','')
		 #print rayon
	       except IndexError:
		    lodg = ''
		   
	       try:
		    sanuzel = grab.doc.select(u'//span[contains(text(),"Санузел")]/following-sibling::span').text().replace(u'не указано','')
	       except IndexError:
		    sanuzel = ''
		     
		     
	       try:
		    okna = grab.doc.select(u'//td[contains(text(),"Вид из окон")]/following-sibling::td').text().replace(u'не указано','')
	       except IndexError:
		    okna = ''
		   
	       try:
	            ln = []
	            for m in grab.doc.select('//ul[@class="geo-tags__list"]/li'):
		         mes = m.text() 
		         ln.append(mes)
	            potolki = ', '.join(ln)
	       except IndexError:
	            potolki =''
		   
	       try:
		    lift = grab.doc.select(u'//td[contains(text(),"Лифт")]/following-sibling::td').text().replace(u'не указано','')
	       except IndexError:
		    lift = ''
		  
	       try:
		    if grab.doc.select(u'//h1[contains(@class,"is-new-building")]').exists() == True:
			 rinok = u'Новостройка'
		    else:
			 rinok = u'Вторичка'
	       except IndexError:
		    rinok = ''
		   
	       try:
		    kons = re.sub(u'^.*(?=консьерж)','', grab.doc.select(u'//*[contains(text(), "консьерж")]').text())[:8].replace(u'консьерж',u'есть')
	       except IndexError:
		    kons = ''
		     
	       try:
		    opis = grab.doc.select(u'//div[@class="foldable-description card-living-content__description"]').text() 
	       except IndexError:
		    opis = ''
		
	       try:
                    phone = re.sub('[^\d\+]','',grab.doc.select(u'//li[@class="offer-card-contacts-phones__item"]/a/@href').text())
	       except IndexError:
		    phone = ''
		   
	       try:
		    lico = grab.doc.select(u'//a[contains(@href,"users")]/span').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//div[@class="offer-card-contacts__person"]/a[contains(@href,"an")]/span').text()
		 #print rayon
	       except IndexError:
		    comp = ''
		    
	       try:
		    data =  re.sub('[^\d\.]','',grab.doc.select(u'//p[contains(text(),"Обновлено")]').text().split(', ')[1])
	       except IndexError:
		    data = ''
		    
	       
		    
	       try:
		    tip_pr = re.sub('[^\d\.]','',grab.doc.select(u'//p[contains(text(),"Обновлено")]').text().split(', ')[0])
		 #print rayon
	       except IndexError:
		    tip_pr = ''
		   
		   
	       
		   
	      
	   
	       projects = {'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
	                   'orentir': orentir,
		           'metro': metro,
		           'potolki': potolki,
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
		           'data':data,
		           'tip_prod':tip_pr
		           
		           }
	     
	     
	     
	       yield Task('write',project=projects,grab=grab)
	     
	     
	     
	     
	     
	     
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
	       print  task.project['object']
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
	       print  task.project['tip_prod']
	       print  task.project['potolki']
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 6,task.project['orentir'])
	       self.ws.write(self.result, 7,task.project['metro'])
	       self.ws.write(self.result, 42,task.project['potolki'])
	       self.ws.write(self.result, 9,task.project['tran'])
	       self.ws.write(self.result, 10,task.project['object'])
	       self.ws.write(self.result, 11,self.oper)
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
	       self.ws.write(self.result, 34, u'N1.RU_Недвижимость')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 41, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 40, task.project['tip_prod'])
	       
	       print('*'*50)
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',dc,'***'
	       print self.oper
	       print('*'*50)
	           
	       self.result+= 1
	       
	       #if self.result > 50:
		    #self.stop()

     
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

     
     

#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
from datetime import datetime
from grab import Grab
import xlsxwriter
import time

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

g = Grab(timeout=20, connect_timeout=200)
i = 0
l= ['http://kvadrat22.ru/mob_sellflatbank-1000-1.html',
    'http://kvadrat24.ru/mob_sellflatbank-1000-1.html',
    'http://kvadrat54.ru/mob_sellflatbank-1000-1.html',
    'http://kvadrat64.ru/mob_sellflatbank-1000-1.html',
    'http://kvadrat66.ru/mob_sellflatbank-1000-1.html',
    'http://kvadrat72.ru/mob_sellflatbank-1000-1.html',
    'http://kvadrat74.ru/mob_sellflatbank-1000-1.html',
    'http://n30.ru/mob_sellflatbank-1000-1.html',
    'http://kemdom.ru/mob_sellflatbank-1000-1.html',
    'http://n002.ru/mob_sellflatbank-1000-1.html',
    'http://kazan-n.ru/mob_sellflatbank-1000-1.html',
    'http://nd27.ru/mob_sellflatbank-1000-1.html',
    'http://nd23.ru/mob_sellflatbank-50-1.html',
    'http://kvadrat22.ru/mob_giveflatbank-1000-1.html',
    'http://kvadrat24.ru/mob_giveflatbank-1000-1.html',
    'http://kvadrat54.ru/mob_giveflatbank-1000-1.html',
    'http://kvadrat64.ru/mob_giveflatbank-1000-1.html',
    'http://kvadrat66.ru/mob_giveflatbank-1000-1.html',
    'http://kvadrat72.ru/mob_giveflatbank-1000-1.html',
    'http://kvadrat74.ru/mob_giveflatbank-1000-1.html',
    'http://n30.ru/mob_giveflatbank-1000-1.html',
    'http://nd23.ru/mob_giveflatbank-50-1.html',
    'http://kemdom.ru/mob_giveflatbank-1000-1.html',
    'http://n002.ru/mob_giveflatbank-1000-1.html',
    'http://kazan-n.ru/mob_giveflatbank-1000-1.html',
    'http://nd27.ru/mob_giveflatbank-1000-1.html']

page = l[i]

while True:  
     class Kvadrat_Kv(Spider):
     
     
     
	  def prepare(self):
	       self.f = page
	       while True:
		    try:
			 time.sleep(1)
			 g.go(self.f)
			 conv = [(u'Хабаровска',u'Хабаровский край'),(u'Барнаула',u'Алтайский край'),
			         (u'Красноярска',u'Красноярский край'),(u'Саратова',u'Саратовская область'),
			         (u'Новосибирска',u'Новосибирская область'),(u'Екатеринбурга',u'Свердловская область'),
			         (u'Тюмени',u'Тюменская область'),(u'Челябинска',u'Челябинская область'),
			         (u'Астрахани',u'Астраханская область'),(u'Кемерово',u'Кемеровская область'),
			         (u'Уфы',u'Башкортостан'),(u'Казани',u'Татарстан'),(u'Краснодара',u'Краснодарский край')]        
			 dt = g.doc.select(u'//a[@class="amenu"][contains(text(),"Недвижимость")]').text().replace(u'Недвижимость','') 
			 self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
			 print self.sub
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 continue
	       self.workbook = xlsxwriter.Workbook(u'kv/Kvadrat_%s' % bot.sub +str(i+1)+ u'_Жилье.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Kvadrat_Жилье')
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
	       
	       self.result= 1
		 
		 
		 
		   
	 
	  def task_generator(self):
               yield Task ('post',url = page,network_try_count=100)



	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//div[@class="dphase"]/following-sibling::a[1]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,network_try_count=100)
	       except DataNotFound:
		    print('*'*100)
		    print '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!','NO PAGE NEXT','!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
		    print('*'*100)
		    logger.debug('%s taskq size' % self.task_queue.size())	
     
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//a[@class="site3"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,network_try_count=100)
	       yield Task("page", grab=grab,network_try_count=100)
		 
	     
	     
	  
	  def task_item(self, grab, task):
	       
	       try:   
	            punkt= grab.doc.select(u'//td[@class="hh"]').text().split(', ')[4].split(' (')[0].replace(u' на карте','')
	       except IndexError:
		    punkt = ''
	       try:
		    ter= grab.doc.select(u'//td[@class="hh"]').text().split(', ')[3].replace(u' на карте','')
	       except IndexError:
		    ter =''
	       try:
		    #uli = re.sub(r'\(.*?$', '',grab.doc.select(u'//td[@class="hh"]').text().split(', ')[3]).replace(u' на карте','')
		    uliza=grab.doc.select(u'//td[@class="hh"]').text().split(', ')[1]
	       except IndexError:
		    uliza = ''
	       try:
		    dom = re.sub('[^\d]','',grab.doc.select(u'//td[@class="hh"]').text().split(', ')[2])
	       except IndexError:
		    dom = ''
	       try:
		    orent = grab.doc.select(u'//span[@class="orientir"]').text().replace(u'(','').replace(u')','')
	       except IndexError:
		    orent = ''
		     
		 
	       try:
		    metro = re.sub('[^\d]','',grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Дата сдачи: ')[1].split(u' год')[0])
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//a[@class="amenu"][contains(text(),"Недвижимость")]').text().replace(u'Недвижимость',u'Недвижимость ')
		 #print rayon
	       except DataNotFound:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//span[@class="object_item_metro_comment"]').text().split(u'мин. ')[1]
	       except IndexError:
		    metro_tr = ''
		    
	       try:
		    tip_ob = u'Квартира' 
	       except DataNotFound:
		    tip_ob = ''
		    
	       try:
		    oper = grab.doc.select(u'//div[@class="a"]').text().split(' ')[0] 
	       except IndexError:
		    oper = ''
		   
	       try:
		    price = grab.doc.select(u'//td[@class="thprice"]').text()   
	       except IndexError:
		   price = ''
		   
	       try:
		    price_m = grab.doc.rex_text(u'Цена за м&sup2;:<br><span class=d>(.*?)</span>').replace('&sup','').replace(';','')
	       except IndexError:
		    price_m = ''
		     
	       try:
		    kol_komnat = re.sub('[^\d]','',grab.doc.select(u'//td[@class="hh"]').text().split(', ')[0])
		
	       except IndexError:
		   kol_komnat = ''
     
	       
     
	       try:
		    plosh_ob = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Общая площадь: ')[1].split(u' м²')[0]+u' м2'
		  #print rayon
	       except IndexError:
		  plosh_ob = ''
     
	       try:
		 plosh_gil = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Жилая площадь: ')[1].split(u' м²')[0]+u' м2'
		  #print rayon
	       except IndexError:
		  plosh_gil = ''
		     
	       try:
		 plosh_kuh = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Площадь кухни: ')[1].split(u' м²')[0]+u' м2'
		  #print rayon
	       except IndexError:
		  plosh_kuh = ''
		  
	       try:
		    plosh_com = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Высота потолков: ')[1].split(u' м')[0].replace(u'потолок ','')+u' м'
	       except IndexError:
		    plosh_com = ''
		    
	       try:
		 et = re.sub('[^\d\/]','',grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Этаж/этажей в доме: этаж ')[1].split(u'Дом')[0]).split('/')[0]
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		 etagn =re.sub('[^\d\/]','',grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Этаж/этажей в доме: этаж ')[1].split(u'Дом')[0]).split('/')[1]
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''
		     
	       try:
		 mat = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Дом(cтроение): ')[1].replace(u'дом ','')[:9]
		 #print rayon
	       except IndexError:
		   mat = '' 
		   
	       #try:
		 #god = grab.doc.select(u'//div[contains(text(),"Год постройки/сдачи:")]/following-sibling::div[@class="propertyValue"]').text()
	       #except DataNotFound:
		   #god = ''
		     
	       try:
		 balkon = grab.doc.select(u'//th[contains(text(),"Балкон:")]/following-sibling::td[contains(text(),"балк")]').text()#.replace(u'нет','')
		 #print rayon
	       except DataNotFound:
		   balkon = ''
		   
	       try:
		 lodg = grab.doc.select(u'//th[contains(text(),"Балкон:")]/following-sibling::td[contains(text(),"лодж")]').text()
		 #print rayon
	       except DataNotFound:
		   lodg = ''
		   
	       try:
		 sanuzel = grab.doc.select(u'//th[contains(text(),"Санузел:")]/following-sibling::td').text().replace(u'нет','')
	       except DataNotFound:
		   sanuzel = ''
		     
		     
	       try:
		 okna = grab.doc.select(u'//th[contains(text(),"Вид из окна:")]/following-sibling::td').text()
	       except DataNotFound:
		   okna = ''
		   
	       
		   
	       try:
		    lift = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Лифт: ')[1].split(u' лифт')[0][:4]
	       except IndexError:
		    lift = ''
		  
	       try:
		 rinok = grab.doc.select(u'//th[contains(text(),"Тип дома:")]/following-sibling::td').text().split(', ')[0]
	       except DataNotFound:
		   rinok = ''
		   
	       try:
		 kons = re.sub(u'^.*(?=консьерж)','', grab.doc.select(u'//*[contains(text(), "консьерж")]').text())[:8].replace(u'консьерж',u'есть')
	       except DataNotFound:
		   kons = ''
		     
	       try:
		  opis = grab.doc.select(u'//div[contains(text(), "Дополнительная информация:")]/span').text() 
	       except IndexError:
		   opis = ''
		
	       try:
		    phone = re.sub('[^\d]','',grab.doc.select(u'//div[@class="divdec"]').text().split(u'телефон:')[1])[:11]
	       except (AttributeError,IndexError):
		   phone = ''
		   
	       try:
		 lico = grab.doc.rex_text(u'Персона для контактов:<br><span class=d>(.*?)</span>')
	       except IndexError:
		   lico = ''
		    
	       try:
		 comp = grab.doc.select(u'//img[@class="thumbnail"][contains(@src,"agency")]/following::b[1]').text()
		 #print rayon
	       except DataNotFound:
		   comp = ''
		    
	       try:
		    data = grab.doc.select(u'//td[@class="tdate"]').text().split(u'создано ')[0].replace(u'обновлено ','').replace('-','.')
	       except DataNotFound:
		    data = ''
		    
	       
		    
	      
		   
		   
	       
		   
	      
	   
	       projects = {'sub': self.sub,
		           'orentir': orent,
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
		           'col_komnat': kol_komnat,
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
		           'phone':phone,
		           'lico':lico,
		           'company':comp,
		           'data':data,
		           'oper':oper
		           }
	     
	     
	     
	       yield Task('write',project=projects,grab=grab)
	     
	     
	     
	     
	     
	     
	  def task_write(self,grab,task):
	       
	       print('*'*50)	       
	       print  task.project['sub']
	       print  task.project['orentir']
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
	       
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 6,task.project['orentir'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 22,task.project['metro'])
	       self.ws.write(self.result, 34,task.project['udall'])
	       self.ws.write(self.result, 9,task.project['tran'])
	       self.ws.write(self.result, 10,task.project['object'])
	       self.ws.write(self.result, 11,task.project['oper'])
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 13, task.project['cena_m'])
	       self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 15, task.project['plosh_ob'])
	       self.ws.write(self.result, 16, task.project['plosh_gil'])
	       self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       self.ws.write(self.result, 29, task.project['plosh_com'])
	       self.ws.write(self.result, 19, task.project['etach'])
	       self.ws.write(self.result, 20, task.project['etashost'])
	       self.ws.write(self.result, 21, task.project['material'])
	       self.ws.write(self.result, 24, task.project['balkon'])
	       self.ws.write(self.result, 25, task.project['logia'])
	       self.ws.write(self.result, 26, task.project['uzel'])
	       self.ws.write(self.result, 27, task.project['okna'])
	       self.ws.write(self.result, 30, task.project['lift'])
	       self.ws.write(self.result, 31, task.project['rinok'])
	       self.ws.write(self.result, 32, task.project['kons'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       #self.ws.write(self.result, 34, u'Брянский сервер недвижимости')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 40, datetime.today().strftime('%d.%m.%Y'))
	       
	       
	       print('*'*50)
	       print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print('*'*50)
	       self.result+= 1
	       
	       
	       #if self.result > 10:
		    #self.stop()

     bot = Kvadrat_Kv(thread_number=5,network_try_limit=2000)
     bot.load_proxylist('/home/oleg/Proxy/tipa.txt','text_file')
     bot.create_grab_instance(timeout=5000, connect_timeout=5000)
     bot.run()
     #print bot.sub
     print(u'Сохранение...')
     print(u'Спим 2 сек...')
     time.sleep(2) 
     bot.workbook.close()
     print('Done!')
     i=i+1
     try:
	  page = l[i]
     except IndexError:
	  break

     
     
#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError,GrabNetworkError,GrabConnectionError,DataNotFound,GrabTooManyRedirectsError
import logging
import base64
from grab import Grab
import re
import time
import math
from datetime import datetime
import xlsxwriter
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logging.basicConfig(level=logging.DEBUG)



i = 0
l= open('Links/Kvartir.txt').read().splitlines()

page = l[i]



while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
           
     class IRR_Kvartir(Spider):
    
    
    
	  def prepare(self):
	             
               self.f = page
               #self.link =l[i]
               #for p in range(1,51):
	            #try:
			 #time.sleep(1)
			 #g = Grab(timeout=20, connect_timeout=20)
			 #g.proxylist.load_file(path='../ivan.txt',proxy_type='http')
			 #g.go(self.f)
			 ##city = [ (u'кой',u'кая'),(u'области',u'область'),(u'ком',u'кий'),
			          ##(u'Москве',u'Москва'),(u'Петербурге',u'Петербург'),
			          ##(u'крае',u'край'),(u'республике ','')]
		    
			 ##dt = g.doc.select(u'//span[@itemprop="name"]').text().replace('Все объявления в ','').replace('Все объявления во ','').replace('/','-')
		    
			 ##self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), city, dt)
			 #try:
                              #self.num= re.sub('[^\d]', '',g.doc.select(u'//div[@class="listingStats"]').text().split('из ')[1])
                         #except IndexError:
	                      #self.num = '1'
			 #self.pag = int(math.ceil(float(int(self.num))/float(30)))
			 #print self.num,self.pag
			 #del g
			 #break
		    #except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError,GrabTooManyRedirectsError):
		         #print g.config['proxy'],'Change proxy'
		         #g.change_proxy()
		         #del g
		         #continue
	       #else:
		    #self.pag = 1	       
               self.workbook = xlsxwriter.Workbook(u'flats/IRR_Жилье_'+str(i+1)+'.xlsx')
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
	       self.ws.write(0, 40, "ДАТА_ОБНОВЛЕНИЯ_ОБЪЯВЛЕНИЯ")
	       self.ws.write(0, 41, "ДАТА_ПАРСИНГА")
	       self.ws.write(0, 42, "МЕСТОПОЛОЖЕНИЕ")
	       self.conv = [(u' августа',u'.08.2019'), (u' июля',u'.07.2019'),
			 (u' мая',u'.05.2019'),(u' июня',u'.06.2019'),
			 (u' марта',u'.03.2019'),(u' апреля',u'.04.2019'),
			 (u' января',u'.01.2019'),(u' декабря',u'.12.2018'),
			 (u' сентября',u'.09.2019'),(u' ноября',u'.11.2018'),
			 (u' февраля',u'.02.2019'),(u' октября',u'.10.2018'), 
			 (u'сегодня,',datetime.today().strftime('%d.%m.%Y'))]	       
	      
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):
	       #for x in range(1,self.pag+1):
                    #link = self.f+'page'+str(x)+'/'
                    #yield Task ('post',url=link.replace(u'page1/',''),refresh_cache=True,network_try_count=100)
	       yield Task ('post',url= self.f,refresh_cache=True,network_try_count=100)
           
           
            
	  def task_post(self,grab,task):
	       
               if grab.doc.select(u'//h4[contains(text(),"Предложения рядом")]').exists()==True:
                    links = grab.doc.select(u'//h4[contains(text(),"Предложения рядом")]/preceding::a[contains(@class,"listing")]')
               else:
                    links = grab.doc.select(u'//a[@class="listing__itemTitle"]')
               for elem in links:
                    ur = grab.make_url_absolute(elem.attr('href'))
                    #print ur
                    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
               yield Task('page', grab=grab,refresh_cache=True,network_try_count=100)  
	       
	       
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
                    ray = grab.doc.select(u'//li[contains(text(),"АО:")]').text().split(': ')[1]
               except IndexError:
                    ray =''
               try:
                    punkt = grab.doc.rex_text(u'address_city":"(.*?)"').decode("unicode_escape")
               except (IndexError,TypeError,ValueError,KeyError):
                    punkt = ''
		    
	       try:
		    ter =  grab.doc.select(u'//li[contains(text(),"Район города:")]').text().split(': ')[1]
	       except IndexError:
		    ter ='' 
	       try:
	            uliza = grab.doc.select(u'//li[contains(text(),"Улица:")]').text().split(': ')[1]
               except IndexError:
		    uliza = ''
		    
               try:
	            dom = grab.doc.select(u'//li[contains(text(),"Дом:")]').text().split(': ')[1].replace('/','|')
	        #print rayon
	       except IndexError:
	            dom = ''
		   
	       
               try:
                    metro = grab.doc.select(u'//li[contains(text(),"Метро:")]').text().split(': ')[1]
               except IndexError:
                    metro = ''
		   
               try:
                    metro_min = grab.doc.select(u'//li[contains(text(),"До метро, минут(пешком):")]').text().split(': ')[1]
               except IndexError:
                    metro_min = ''
		   
               try:
		    try:
                         price = grab.doc.select(u'//div[@class="productPage__price js-contentPrice"]').text()
		    except IndexError:
			 price = grab.doc.select(u'//div[@class="productPage__price"]').text()
               except IndexError:
                    price = ''
		   
               try:
                    kol_komnat = grab.doc.select(u'//li[contains(text(),"Комнат в квартире:")]').number()
                #print rayon
               except IndexError:
                    kol_komnat = ''
   
               try:
                    plosh_ob = grab.doc.select(u'//li[contains(text(),"Общая площадь:")]').text().split(': ')[1]
                  #print rayon
               except IndexError:
                    plosh_ob = ''
     
               try:
                    plosh_gil = grab.doc.select(u'//li[contains(text(),"Жилая площадь:")]').text().split(': ')[1]
                  #print rayon
               except IndexError:
                    plosh_gil = ''
		   
               try:
                    plosh_kuh = grab.doc.select(u'//li[contains(text(),"Площадь кухни:")]').text().split(': ')[1]
                  #print rayon
               except IndexError:
                    plosh_kuh = ''
		    
               try:
                    et = grab.doc.select(u'//li[contains(text(),"Этаж:")]').number()
                 #print price + u' руб'	    
               except IndexError:
                    et = '' 
		   
               try:
                    etagn = grab.doc.select(u'//li[contains(text(),"Этажей в здании:")]').number()
                 #print price + u' руб'	    
               except IndexError:
                    etagn = ''
		   
               try:
                    mat = grab.doc.select(u'//li[contains(text(),"Материал стен:")]').text().split(': ')[1]
                 #print rayon
               except IndexError:
                    mat = '' 
		   
               try:
                    god = grab.doc.select(u'//li[contains(text(),"Год постройки:")]').text().split(': ')[1]
               except IndexError:
                    god = ''
		   
               try:
                    balkon = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Балкон/Лоджия")]').text().split('/')[0].replace(u'Балкон',u'есть')
                 #print rayon
               except IndexError:
                    balkon = ''
		   
               try:
                    lodg = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Балкон/Лоджия")]').text().split('/')[1].replace(u'Лоджия',u'есть')
                 #print rayon
               except IndexError:
                    lodg = ''
		   
               try:
                    sanuzel = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Санузел:")]').text().split(': ')[1]
               except IndexError:
                    sanuzel = ''
               try:
                    sost = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Ремонт:")]').text().split(': ')[1]
               except IndexError:
                    sost = ''
		   
               try:
                    potolki = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Высота потолков:")]').text().split(': ')[1]
               except IndexError:
                    potolki = ''
		   
               try:
                    lift = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Лифты в здании")]').text().replace(u'Лифты в здании',u'есть')
               except IndexError:
                    lift = ''
		  
               try:
                    try:
	                 rinok = grab.doc.select(u'//span[@itemprop="name"][contains(text(),"Новостройки")]').text()
                    except IndexError:
	                 rinok = grab.doc.select(u'//span[@itemprop="name"][contains(text(),"Вторичный рынок")]').text()
               except IndexError:
	            rinok = ''
		   
               try:
                    kons = re.sub(u'^.*(?=консьерж)','', grab.doc.select(u'//*[contains(text(), "консьерж")]').text())[:8].replace(u'консьерж',u'есть')
               except IndexError:
                    kons = ''
		   
               try:
                    opis = grab.doc.select(u'//h4[contains(text(),"Описание")]/following-sibling::p').text() 
               except IndexError:
                    opis = ''
		   
               try:
                    phone = re.sub('[^\d]', '',base64.b64decode(grab.doc.select('//input[@name="phoneBase64"]').attr('value')))
               except (AttributeError,IndexError):
                    phone = ''
		   
               try:
                    lico = grab.doc.select(u'//div[@class="productPage__infoTextBold productPage__infoTextBold_inline"]').text()
               except IndexError:
                    lico = ''
		   
               try:
                    comp = grab.doc.select(u'//div[@class="productPage__infoTextBold productPage__infoTextBold_inline"]/a[contains(@href,"russia")]').text()
                 #print rayon
               except IndexError:
                    comp = ''
		   
               try:
		    data =grab.doc.rex_text(u'date_create":"(.*?)"}').split(' ')[0].replace('-','.')
               except IndexError:
                    data = ''
		    
	       try:
		    d1 = grab.doc.select(u'//div[@class="productPage__createDate"]').text()
	            data1 = reduce(lambda d1, r: d1.replace(r[0], r[1]), self.conv, d1).replace(u'Размещено ','')
	       except IndexError:
	            data1 = ''		    
               try:
                    ohrana = grab.doc.select(u'//i[@class="icon icon_spot"]/following-sibling::div[1]').text()
               except IndexError:
                    ohrana = ''
		    
	    
	       if 'sale' in task.url:
		    oper = u'Продажа'
	       elif 'rent' in task.url:
		    oper = u'Аренда'
	       else:
		    oper =''
	      
		  
	       projects = {'sub': sub,
		           'rayon': ray,
		           'punkt': punkt,
	                   'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
	                   'metro': metro,
	                   'udall': metro_min,
	                   'cena': price,
		           'col_komnat': kol_komnat,
		           'plosh_ob':plosh_ob,
		           'plosh_gil': plosh_gil,
		           'plosh_kuh': plosh_kuh,
		           'etach': et,
		           'etashost': etagn,
		           'material': mat,
		           'god_post': god,
		           'balkon': balkon,
		           'logia': lodg,
	                   'uzel':sanuzel,
		           'sost': sost,
		           'vis_potolok':potolki,
		           'lift':lift,
		           'rinok': rinok,
		           'kons':kons,
	                   'operacia':oper,
		           'opis':opis,
		           'url':task.url,
	                   'phone':phone,
		           'lico':lico.replace(comp,''),
		           'company':comp,
	                   'data':data,
	                   'data1':data1[:10],
	                   'ochrana':ohrana}
	
	
	
	       yield Task('write',project=projects,grab=grab)
	
	
	
	
	
	
          def task_write(self,grab,task):
	       
               print('*'*50)	       
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['metro']
	       print  task.project['udall']
	       print  task.project['cena']
	       print  task.project['col_komnat']
	       print  task.project['plosh_ob']
	       print  task.project['plosh_gil']
	       print  task.project['plosh_kuh']
	       print  task.project['etach']
	       print  task.project['etashost']
	       print  task.project['material']
	       print  task.project['god_post']
	       print  task.project['balkon']
	       print  task.project['logia']
	       print  task.project['uzel']
	       print  task.project['sost']
	       print  task.project['vis_potolok']
	       print  task.project['lift']
	       print  task.project['rinok']
	       print  task.project['kons']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['data1']
               print  task.project['ochrana']
              
	 
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 3, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 7, task.project['metro'])
	       self.ws.write(self.result, 8, task.project['udall'])
	       self.ws.write(self.result, 10, 'Квартира')
	       self.ws.write(self.result, 11, task.project['operacia'])
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 15, task.project['plosh_ob'])
	       self.ws.write(self.result, 16, task.project['plosh_gil'])
	       self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       self.ws.write(self.result, 19, task.project['etach'])
	       self.ws.write(self.result, 20, task.project['etashost'])
	       self.ws.write(self.result, 21, task.project['material'])
	       self.ws.write(self.result, 22, task.project['god_post'])
	       self.ws.write(self.result, 24, task.project['balkon'])
	       self.ws.write(self.result, 25, task.project['logia'])
	       self.ws.write(self.result, 26, task.project['uzel'])
	       self.ws.write(self.result, 28, task.project['sost'])
	       self.ws.write(self.result, 29, task.project['vis_potolok'])
	       self.ws.write(self.result, 30, task.project['lift'])
	       self.ws.write(self.result, 31, task.project['rinok'])
	       self.ws.write(self.result, 32, task.project['kons'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, 'Из рук в руки')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 40, task.project['data1'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 41, datetime.today().strftime('%d.%m.%Y'))
               self.ws.write(self.result, 42, task.project['ochrana'])
               
	       print('*'*50)
	       print 'Ready - '+str(self.result)
	       print 'Tasks - %s' % self.task_queue.size()
	       print '*',i+1,'/',len(l),'*'
	       print  task.project['operacia']
	       print('*'*50)
	       self.result+= 1
	       
	       
	       #if self.result > 10:
		    #self.stop()
               #if str(self.result) == str(self.num):
	            #self.stop()		    


     bot = IRR_Kvartir(thread_number=5,network_try_limit=1000)
     #bot.setup_queue('mongo', database='IrrFlat',host='192.168.10.200')
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     #command = 'mount -a'
     #p = os.system('echo %s|sudo -S %s' % ('1122', command))
     #print p
     #time.sleep(2)
     bot.workbook.close()
     #workbook.close()
     print('Done!') 
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
     
time.sleep(5)
os.system("/home/oleg/pars/irr/zag.py")
#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError,GrabTooManyRedirectsError
import logging
import random
from grab import Grab
import re
import time
import math
from datetime import datetime
import xlsxwriter
import base64
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logging.basicConfig(level=logging.DEBUG)



i = 0
l= open('Links/Zemm.txt').read().splitlines()

page = l[i]



while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
           
     class IRR_Kvartir(Spider):
    
	  def prepare(self):
	       self.f = page.replace(u'zemelniy-uchastok','kvartiru')
	       print self.f
	       for p in range(1,31):
		    try:
			 time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http') 
			 g.go(self.f)
			 print g.doc.code
			 self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="search-result"]/span').text())
		         self.pag = int(math.ceil(float(int(self.num))/float(30)))
			 print self.pag,self.num
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,ValueError,DataNotFound,GrabTooManyRedirectsError, GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.pag = 1
		    self.num = 1    
	       
               self.workbook = xlsxwriter.Workbook(u'flats/Gdeetotdom_Flats'+str(i+1)+'.xlsx')
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
	       self.conv = [(u' августа',u'.08.'), (u' июля',u'.07.'),
			 (u' мая',u'.05.'),(u' июня',u'.06.'),
			 (u' марта',u'.03.'),(u' апреля',u'.04.2018'),
			 (u' января',u'.01.'),(u' декабря',u'.12.'),
			 (u' сентября',u'.09.'),(u' ноября',u'.11.'),
			 (u' февраля',u'.02.'),(u' октября',u'.10.'), 
			 (u'сегодня,',datetime.today().strftime('%d.%m.%Y'))]	       
	      
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):	     
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'?page=%d'% x,refresh_cache=True,network_try_count=10)

	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//a[@class="b-objects-list__control-btn"][contains(text(),"Узнать больше")]'):
                    ur = base64.b64decode(elem.attr('data-hide'))  
                    #print ur
                    yield Task('item', url=ur,refresh_cache=True,network_try_count=10)

        
	  def task_item(self, grab, task):
	       try:
                    ray = grab.doc.select(u'//span[contains(text(), "Округ")]/following::div[1]/div/a').text()
               except IndexError:
                    ray =''
               try:
                    punkt = grab.doc.select(u'//span[contains(text(), "Населённый пункт")]/following::div[1]/div/a').text()
               except (IndexError,TypeError,ValueError,KeyError):
                    punkt = ''
		    
	       try:
		    ter =  grab.doc.select(u'//span[contains(text(), "Район")]/following::div[1]/div/a').text()
	       except IndexError:
		    ter ='' 
	       try:
		    try:
	                 uliza = grab.doc.select(u'//span[contains(text(), "Улица")]/following::div[1]/div/a').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//div[@class="title"][contains(text(), "Информация о доме")]').text().split(u', ')[0].replace(u'Информация о доме ', '')
               except IndexError:
		    uliza = ''
		    
               try:
	            dom = grab.doc.select(u'//span[contains(text(), "Здание")]/following::div[1]/div').text()
	       except IndexError:
	            dom = ''
		   
	       
               try:
                    metro = grab.doc.select(u'//div[@class="address-params__metro"]/a[@class="linking"]').text()
               except IndexError:
                    metro = ''
		   
               try:
                    metro_min = grab.doc.select(u'//div[@class="address-params__metro"]/em').text().split(u' минут ')[0]
               except IndexError:
                    metro_min = ''
		   
               try:
		    price = grab.doc.select(u'//ul[@class="price-block"]/li[1]').text().replace(u'Цена ', '')
               except IndexError:
                    price = ''
		   
               try:
                    kol_komnat = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Комнат")]/following::div[@class="b-dotted-block__right"][1]/span').number()
               except IndexError:
                    kol_komnat = ''
   
               try:
                    plosh_ob = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Общая площадь")]/following::div[@class="b-dotted-block__right"][1]/span').text()
               except IndexError:
                    plosh_ob = ''
     
               try:
                    plosh_gil = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Жилая площадь")]/following::div[@class="b-dotted-block__right"][1]/span').text()
               except IndexError:
                    plosh_gil = ''
		   
               try:
                    plosh_kuh = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Площадь кухни")]/following::div[@class="b-dotted-block__right"][1]/span').text()
                  #print rayon
               except IndexError:
                    plosh_kuh = ''
		    
               try:
                    et = grab.doc.select(u'//span[contains(text(),"Этаж")]/following::span[@class="b-dotted-block__inner"][contains(text(),"из")]').text().split(u' из ')[0]
               except IndexError:
                    et = '' 
		   
               try:
                    etagn = grab.doc.select(u'//span[contains(text(),"Этажность")]/following::span[1]').text()
               except IndexError:
                    etagn = ''
		   
               try:
                    mat = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Тип строения")]/following::div[@class="b-dotted-block__right"][1]/span').text()
                 #print rayon
               except IndexError:
                    mat = '' 
		   
               try:
                    god = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Год постройки")]/following::div[@class="b-dotted-block__right"][1]/span').text()
               except IndexError:
                    god = ''
		   
               try:
                    balkon = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Балкон/лоджия")]/following::div[@class="b-dotted-block__right"][1]/span').text()
                 #print rayon
               except IndexError:
                    balkon = ''
		   
               try:
                    lodg = grab.doc.select(u'//ul[@class="price-block"]/li[2]').text()
               except IndexError:
                    lodg = ''
		   
               try:
                    sanuzel = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Санузел")]/following::div[@class="b-dotted-block__right"][1]/span').text()
               except IndexError:
                    sanuzel = ''
               try:
                    sost = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Ремонт")]/following::div[@class="b-dotted-block__right"][1]/span').text()
               except IndexError:
                    sost = ''
		   
               try:
                    potolki = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Площадь комнат")]/following::div[@class="b-dotted-block__right"][1]/span').text()
               except IndexError:
                    potolki = ''
		   
               try:
                    lift = grab.doc.select(u'//span[@class="b-dotted-block__inner"][contains(text(), "Лифт")]/following::div[@class="b-dotted-block__right"][1]/span').text()
               except IndexError:
                    lift = ''
		  
               try:
                    rinok = grab.doc.select(u'//div[@class="address-params__metro"]/em').text().split(u' минут ')[1]
               except IndexError:
	            rinok = ''
		   
               try:
                    kons = grab.doc.select(u'//span[contains(text(), "Регион")]/following::div[1]/div/a').text()
               except IndexError:
                    kons = ''
		   
               try:
                    opis = grab.doc.select(u'//div[@class="description"]').text() 
               except IndexError:
                    opis = ''
		   
               url1 = re.sub('[^\d]','',task.url)
	       try:
		    phone_url = task.url.split(u'ru')[0]+u'ru'+'/classifiedAjax/showPhones/'+url1+'/print/?type=agent'    
		    headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
			      'Accept-Encoding': 'gzip,deflate',
			      'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			      'Cookie': 'sessid='+url1+'.'+url1,
			      'Host': 'www.gdeetotdom.ru',
			      'Referer': task.url,
			      'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:23.0) Gecko/20131011 Firefox/23.0',
			      'X-Requested-With': 'XMLHttpRequest'}
		    g2 = grab.clone(headers=headers,proxy_auto_change=True)
		    g2.request(post=[('type','agent')],headers=headers,url=phone_url) 
		    phone =  re.sub('[^\d\+\,]','',re.findall('phones(.*?),"statlink',g2.doc.body)[0]) 
		    print 'Phone-OK'
		    del g2
	       except (IndexError,GrabConnectionError,GrabNetworkError,GrabTimeoutError):
	            phone = random.choice(list(open('../phone.txt').read().splitlines()))
		   
               try:
                    lico = grab.doc.select(u'//div[@class="realtor-contacts__name"]/span').text()
               except IndexError:
                    lico = ''
		   
               try:
                    comp = grab.doc.select(u'//div[@class="realtor-contacts__name"]/p').text()
                 #print rayon
               except IndexError:
                    comp = ''
		   
               try:
		    data = grab.doc.select(u'//li[@class="activity__publish"]').text().replace(u'Опубликовано ', '').replace(u' г.','')
               except IndexError:
                    data = ''
		    
	       try:
	            data1 = grab.doc.select(u'//li[@class="activity__update"]').text().replace(u'Обновлено ', '').replace(u' г.','')
	       except IndexError:
	            data1 = ''		    
               try:
                    ohrana = grab.doc.select(u'//div[@class="address-line"]').text()
               except IndexError:
                    ohrana = ''
		    
	    
	       try:
                    oper = grab.doc.select(u'//h1').text().split(' ')[0]
               except IndexError:
	  	    oper =''
	      
	       data1 = reduce(lambda data1, r: data1.replace(r[0], r[1]), self.conv, data1)
	       data = reduce(lambda data, r: data.replace(r[0], r[1]), self.conv, data)
	       
	       if kons == u"Москва":
		    punkt= u"Москва"
	       elif kons == u"Санкт-Петербург":
		    punkt='Санкт-Петербург'
	       else:
		    punkt = punkt
		
	       projects = {'rayon': ray,
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
		           'etach': et.replace(etagn,''),
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
		           'lico':lico,
		           'company':comp,
	                   'data':data,
	                   'data1':data1,
	                   'ochrana':ohrana}
	
	
	
	       yield Task('write',project=projects,grab=grab)
	
	
	
	
	
	
          def task_write(self,grab,task):
	       
               print('*'*50)	       
	       print  task.project['kons']
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
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['data1']
               print  task.project['ochrana']
              
	 
	       self.ws.write(self.result, 0, task.project['kons'])
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
	       self.ws.write(self.result, 13, task.project['logia'])
	       self.ws.write(self.result, 26, task.project['uzel'])
	       self.ws.write(self.result, 28, task.project['sost'])
	       self.ws.write(self.result, 18, task.project['vis_potolok'])
	       self.ws.write(self.result, 30, task.project['lift'])
	       self.ws.write(self.result, 9, task.project['rinok'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, 'ГДЕЭТОТДОМ.РУ')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 40, task.project['data1'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 41, datetime.today().strftime('%d.%m.%Y'))
               self.ws.write(self.result, 42, task.project['ochrana'])
               
	       print('*'*20)
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       print '*',i+1,'/',len(l),'*'
	       print  task.project['operacia']
	       print('*'*20)
	       self.result+= 1
	       
	       #if str(self.result) == str(self.num):
		    #self.stop()	
		    
	       if self.result > 6000:
	            self.stop()	       
	    


     bot = IRR_Kvartir(thread_number=15,network_try_limit=100)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=500, connect_timeout=500)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     time.sleep(2)
     bot.workbook.close()
     print('Done!') 
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
     

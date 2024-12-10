#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError,GrabTooManyRedirectsError
import logging
from grab import Grab
import re
import os
import math
from head import agents
import random
import base64
import time
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)





i = 0
l= open('Links/Comm.txt').read().splitlines()
dc = len(l)
page = l[i]


while i < dc:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class Gdedom_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       if 'kupit' in page:
                    self.oper = u'Продажа' 
               elif 'snyat' in page:
	            self.oper = u'Аренда'
	       for p in range(1,31):
                    try:
                         time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http') 
                         g.go(self.f)
			 print g.doc.code
			 self.sub = g.doc.select(u'//li[@class="ssp-breadcrumbs-item last"]/span').text().split(u' в ')[1].replace(u'регионе ','').replace(u'Подмосковье',u'Московская область')
			 self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="search-result"]/span').text())
			 try:
			      self.pag = int(math.ceil(float(int(self.num))/float(30)))
			 except ValueError:
			      self.pag = 0
			      self.num = 0
			      self.stop()
			 if self.pag > 50:
			      self.pag = 50
			      self.num = 1500
			 else:
			      self.pag = self.pag
		              self.num = self.num
			      
			 print self.sub,self.oper,self.pag,self.num
			 del g
			 break
                    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabTooManyRedirectsError, GrabConnectionError):
                         print g.config['proxy'],'Change proxy'
                         g.change_proxy()
			 del g
                         continue
                    
	       else:
	            self.sub = ''
		    self.pag = 1
		    self.stop()

                    
	       self.workbook = xlsxwriter.Workbook(u'com/Gdeetotdom_Коммерческая_'+self.oper+str(i+1)+'.xlsx')
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
	       self.ws.write(0, 36, u"ТРАССА")
	       self.ws.write(0, 37, u"ПАРКОВКА")    
	       self.result= 1
	       self.g = 0
    
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
		    ray = grab.doc.select(u'//li[@class="ssp-linking-item"]/a[contains(text(),"районе")]').text().split(u'районе ')[1]
	       except IndexError:
	            ray = '' 
	       try:
		    if self.sub == u"Москве":
	                 punkt= u"Москва"
		    elif self.sub == u"Санкт-Петербурге":
			 punkt='Санкт-Петербург'
		    else:
			 punkt = grab.doc.select(u'//div[@class="address-line"]').text().split(u', ')[1]
	       except IndexError:
		    punkt = ''
		    
	       try:
		    try:
		         ter= grab.doc.rex_text(u'Продажа (.*?) в')
		    except IndexError:
			 ter= grab.doc.rex_text(u'Аренда (.*?) в')
	       except IndexError:
		    ter =''
		    
	       try:
		    uliza = grab.doc.select(u'//span[contains(text(),"Класс здания")]/following::span[1]').text()
	       except IndexError:
	            uliza = ''
		    
	       try:
		    dom = grab.doc.select(u'//span[contains(text(),"Год постройки")]/following::span[1]').number()
	       except IndexError:
		    dom = ''
		    
	       try:
		    trassa = grab.doc.select(u'//h1').text()
		     #print rayon
	       except IndexError:
		    trassa = ''
		    
	       try:
		    udal = grab.doc.select(u'//span[contains(text(),"Тип строения")]/following::span[1]').text()
	       except IndexError:
		    udal = ''
		    
	       try:
		    price = grab.doc.select(u'//ul[@class="price-block"]/li[1]').text().replace(u'Цена ', '').replace(u'i',u'руб')
	       except IndexError:
		    price = ''
		    
	       
		    
	       try:
		    plosh = grab.doc.select(u'//span[contains(text(),"Общая площадь")]/following::span[1]').text()
	       except IndexError:
		    plosh = ''

	       try:
		    vid = grab.doc.select(u'//div[@class="address-line"]').text()
	       except IndexError:
		    vid = '' 
		    
		    
	       try:
		    ohrana = grab.doc.select(u'//a[@class="ssp-breadcrumbs-link"][contains(@href,"metro")]/span').text()
	       except IndexError:
		    ohrana =''
	       try:
		    gaz = grab.doc.select(u'//div[@class="address-line"]').text()
	       except IndexError:
		    gaz =''
	       try:
		    voda = grab.doc.select(u'//span[contains(text(),"Метро")]/following::span/em').text()
	       except IndexError:
		    voda =''
	       try:
		    kanal = grab.doc.select(u'//span[contains(text(),"Этаж")]/following::span[@class="b-dotted-block__inner"][contains(text(),"из")]').text().split(u' из ')[0]
	       except IndexError:
		    kanal =''
	       try:
		    elek = re.sub('[^\d\.]','',grab.doc.rex_text(u'NearestLatlng = (.*?)]').split(', ')[0])
	       except IndexError:
		    elek =''
		    
	       try:
		    lng = re.sub('[^\d\.]','',grab.doc.rex_text(u'NearestLatlng = (.*?)]').split(', ')[1])
	       except IndexError:
	            lng =''		    
	       try:
		    teplo = grab.doc.select(u'//span[contains(text(),"Этажность")]/following::span[1]').text()
	       except IndexError:
		    teplo =''
		    
	      
	       try:
		    opis = grab.doc.select(u'//div[@class="description"]').text() 
	       except IndexError:
		    opis = ''
		    
	       try:
	            park = grab.doc.select(u'//span[contains(text(),"Парковка")]/following::span[1]').text() 
	       except IndexError:
		    park = ''		    
		    
	       
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
		    
		    
			 
	       #if self.g == 76:
		    #self.g = 0
	       #else:
		    #self.g+= 1		    
              
     
	       try:
		    lico = grab.doc.select(u'//div[@class="realtor-contacts__name"]/span').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//div[@class="realtor-contacts__name"]/p').text()#.replace(u'Агентство ','')
	       except IndexError:
		    comp = ''
		    
	       try:
	            conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
		         (u' мая ',u'.05.'),(u' июня ',u'.06.'),
		         (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
		         (u' января ',u'.01.'),(u' декабря ',u'.12.'),
		         (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
		         (u' февраля ',u'.02.'),(u' октября ',u'.10.')] 
		    d = grab.doc.select(u'//li[@class="activity__publish"]').text().replace(u'Опубликовано ', '').replace(u' г.','')
		    d1 = grab.doc.select(u'//li[@class="activity__update"]').text().replace(u'Обновлено ', '').replace(u' г.','')
		    data = reduce(lambda d, r: d.replace(r[0], r[1]), conv, d)
		    data1 = reduce(lambda d1, r: d1.replace(r[0], r[1]), conv, d1)
	       except DataNotFound:   
	            data = ''
		    data1 = ''
			 
	       
	       clearText = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", opis)
	       clearText = re.sub(u"[.,\-\s]{3,}", " ", clearText)
		    
	       projects = {'url': task.url,
		           'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
		           'trassa': trassa,
		           'udal': udal,
		           'cena': price,
		           'plosh':plosh,
		           'vid': vid,
		           'ohrana':ohrana,
		           'gaz': gaz,
		           'voda': voda,
		           'kanaliz': kanal.replace(teplo,''),
		           'electr': elek,
		           'teplo': teplo,
	                   'dol': lng,
		           'opis':clearText,
		           'phone':phone,
	                   'parkov':park,
		           'lico':lico,
	                   'data1':data1,
		           'company':comp,
		           'data':data}
	       
	       yield Task('write',project=projects,grab=grab)
		 
	  def task_write(self,grab,task):
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['trassa']
	       print  task.project['udal']
	       print  task.project['cena']
	       print  task.project['plosh']
	       print  task.project['ohrana']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['dol']
	       print  task.project['teplo']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['data1']
	       print  task.project['gaz']
	       
	       
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 3, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 7, task.project['teritor'])
	       self.ws.write(self.result, 10, task.project['ulica'])
	       self.ws.write(self.result, 17, task.project['dom'])
	       self.ws.write(self.result, 33, task.project['trassa'])
	       self.ws.write(self.result, 8, task.project['udal'])
	       self.ws.write(self.result, 28, self.oper)
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 24, task.project['vid'])
	       #self.ws.write(self.result, 36, task.project['gaz'])
	       self.ws.write(self.result, 27, task.project['voda'])
	       self.ws.write(self.result, 15, task.project['kanaliz'])
	       self.ws.write(self.result, 34, task.project['electr'])
	       self.ws.write(self.result, 35, task.project['dol'])
	       self.ws.write(self.result, 37, task.project['parkov'])
	       self.ws.write(self.result, 16, task.project['teplo'])
	       self.ws.write(self.result, 26, task.project['ohrana'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'ГдеЭтотДом.Ру')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 30, task.project['data1'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       #print task.sub
	       
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       print '*',i+1,'/',dc,'*'
	       print self.oper
	       print  task.project['vid']
	       #print agents[self.g]
	       print('*'*50) 
	       self.result+= 1
		    
		    
		    
	       #if self.result >= 10:
		    #self.stop()
	       #if int(self.result) >= int(self.num)-3:
	            #self.stop()		    
     
	  
     bot = Gdedom_Zem(thread_number=5,network_try_limit=100)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=500, connect_timeout=500)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     bot.workbook.close()
     print('Done')
     i=i+1
     try:
          page = l[i]
     except IndexError:
	  break
     







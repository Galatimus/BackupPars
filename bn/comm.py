#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import os
import math
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')



logging.basicConfig(level=logging.DEBUG)






i = 0
l= open('links/com_arenda.txt').read().splitlines()
page = l[i]
oper = u'Аренда'




while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class Bn_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	      
               while True:
                    try:
                         time.sleep(2)
			 g = Grab(timeout=10, connect_timeout=10)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')			 
                         g.go(self.f)
			 self.sub = g.doc.select(u'//a[@id="region_btn"]').text().replace('/','-')
			 if g.doc.select(u'//div[@class="count"]').exists()==True:
			      self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="count"]').text())
			      self.pag = int(math.ceil(float(int(self.num))/float(50)))
			 else:
			      self.num = 0
			      self.pag = 0
			 print self.sub,self.pag,self.num
			 del g
			 break
			 
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
		    
	       self.workbook = xlsxwriter.Workbook(u'com/Bn_%s' % bot.sub + u'_Коммерческая_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'bn_Коммерческая')
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
	       self.result= 1
	       
                
                
                
                
	  def task_generator(self):
	       if self.pag == 0:
		    self.stop()
	       else:
		    for x in range(0,self.pag+1):
                         yield Task ('post',url=self.f+'?start='+str(x*50),refresh_cache=True,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@class="underline"][@onclick="return false;"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	      
            
	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//ul[@id="breadcrumb"]/li/a[contains(text(),"район")]').text()
	       except IndexError:
	            mesto =''
	       try:
		    mesto1 = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text().replace(u'Объект на карте','')
	       except IndexError:
	            mesto1 =''		    
	       try:
		    if self.sub == u"Москва":
		         punkt= u"Москва"
		    elif self.sub == u"Санкт-Петербург":
		         punkt= u"Санкт-Петербург"
		    elif self.sub == u"Севастополь":
		         punkt= u"Севастополь"
		    else:   
	                 punkt = grab.doc.select(u'//h1').text().split(', ')[1]
	       except IndexError:
	            punkt = ''	       
		
               try:
                    ter= grab.doc.select(u'//ul[@id="breadcrumb"]/li/a[contains(text(),"АО")]').text()
               except IndexError:
                    ter =''
	         
               try:
		    if 'offices' in task.url:
			 tip = u'Офисный'
		    elif 'service' in task.url:
			 tip = u'Сферауслуг'
		    elif 'different' in task.url:
			 tip = u'ПСН'
		    elif 'storage' in task.url:
			 tip = u'производственно-складской'
		    else:
			 tip =''
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//h1').text().split(', ')[0].replace(u'Продажа ','').replace(u'Аренда ','').replace(u'офиса',u'офис')
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//td[contains(text(),"Тип здания")]/following-sibling::td').text()
               except IndexError:
                    klass = ''
               try:
		    #try:
                    price = grab.doc.select(u'//dt[contains(text(),"Цена")]/following-sibling::dd').text()
		    #except IndexError:
			 #price = grab.doc.select(u'//span[@id="price-total-0"]').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//dt[contains(text(),"Общая площадь")]/following-sibling::dd').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = re.sub('[^\d]','',grab.doc.select(u'//dt[contains(text(),"Этаж")]/following-sibling::dd').text().split(u' в ')[0])
               except IndexError:
                    ohrana =''
               try:
                    gaz =  re.sub('[^\d]','',grab.doc.select(u'//dt[contains(text(),"Этаж")]/following-sibling::dd').text().split(u' в ')[1])
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//td[contains(text(),"Год постройки")]/following-sibling::td').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//h1').text()
               except IndexError:
                    kanal =''
               try:
		    if 'freestanding' in task.url:
			 elek = u'ОСЗ'
		    else:
			 elek =''
                    #elek = grab.doc.select(u'//td[contains(text(),"Электричество")]/following-sibling::td').text()
               except IndexError:
                    elek =''
               try:
                    teplo = grab.doc.select(u'//td[contains(text(),"Водоснабжение")]/following-sibling::td').text()
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//div[@id="description"]').text() 
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//dt[contains(text(),"Метро")]/following-sibling::dd').text()
	       except IndexError:
                    lico = ''
               try:
		    try:
                         comp = grab.doc.select(u'//dt[contains(text(),"Продает")]/following-sibling::dd').text()
		    except IndexError:
			 comp = grab.doc.select(u'//dt[contains(text(),"Сдает")]/following-sibling::dd').text()
               except IndexError:
                    comp = ''
               
	       try:    
	            data = grab.doc.select(u'//dt[contains(text(),"Дата обновления")]/following-sibling::dd').text()#.split(u' в ')[0]
		     #= reduce(lambda d, r: d.replace(r[0], r[1]), conv, d)
	       except IndexError:
		    data =''
	       try:
		    data1 = grab.doc.select(u'//dt[contains(text(),"Дата размещения")]/following-sibling::dd').text()#.split(u' в ')[0].replace(u'обновлено ','')
		     #= reduce(lambda d1, r: d1.replace(r[0], r[1]), conv, d1)
	       except IndexError:
		    data1=''
	       
	       try:
                    phone = grab.doc.select(u'//dt[contains(text(),"Телефон")]/following-sibling::dd').text()
               except IndexError:
	            phone = ''
          
	       
		    
     
	
               projects = {'sub': self.sub,
	                  'adress': mesto,
	                  'adress1': mesto1,
	                   'terit':ter, 
	                   'punkt':punkt, 
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
	                   'lico':lico,
	                   'company': comp,
	                   'data':data,
	                   'data1':data1}
	       
	       
	       try:
		    #ad= re.findall(u"full_address='(.*?)';$",grab.doc.select(u'//div[@id="map"]/following-sibling::script').text())[0]
		    #print ad
		    link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+self.sub+', '+mesto1
		    yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    yield Task('adres',grab=grab,project=projects)	       
	  
	  
	  def task_adres(self, grab, task):
	       try:
		    uliza=grab.doc.rex_text(u'ThoroughfareName":"(.*?)"')
	       except IndexError:
		    uliza = ''
	       try:
		    dom = grab.doc.rex_text(u'PremiseNumber":"(.*?)"')
	       except IndexError:
		    dom = ''
	  
	       project2 ={'ulica':uliza,
	                  'dom':dom}	  
	  
	       yield Task('write',project=task.project,proj=project2,grab=grab)
	  
	  
	  
	  def task_write(self,grab,task):
	       #time.sleep(1)
	       print('*'*100)	       
	       print  task.project['sub']
	       print  task.project['punkt']
	       print  task.project['adress']
	       print  task.project['adress1']
	       print  task.project['terit']
	       print  task.proj['ulica']
	       print  task.proj['dom']
	       print  task.project['tip']
	       print  task.project['naz']
	       print  task.project['klass']
	       print  task.project['cena']
	       print  task.project['plosh']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['teplo']
	       print  task.project['ohrana']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['data1']
	      
	  
	       
	       
     
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['adress'])
	       self.ws.write(self.result, 3, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.proj['ulica'])
	       self.ws.write(self.result, 5, task.proj['dom'])
	       self.ws.write(self.result, 7, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       self.ws.write(self.result, 16, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 15, task.project['ohrana'])
	       self.ws.write(self.result, 16, task.project['gaz'])
	       self.ws.write(self.result, 17, task.project['voda'])
	       self.ws.write(self.result, 33, task.project['kanaliz'])
	       self.ws.write(self.result, 8, task.project['electr'])
	       #self.ws.write(self.result, 21, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'Система "Бюллетень Недвижимости"')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 30, task.project['data'])
	       self.ws.write(self.result, 29, task.project['data1'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 28, oper)
	       self.ws.write(self.result, 24, task.project['adress1'])
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+self.num
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*100)
	       self.result+= 1
	       
	      
	       
	       
	       
	       #if self.result >= 10:
	            #self.stop()	       


     bot = Bn_Com(thread_number=5, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     bot.workbook.close()
     print('Done!')
     i=i+1
     try:
          page = l[i]
     except IndexError:
          if oper == u'Аренда':
               i = 0
               l= open('links/com_prod.txt').read().splitlines()
               page = l[i]
               oper = u'Продажа'
          else:
               break
       
     
     

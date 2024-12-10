#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
import math
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import time
import os
#from sub import conv
from datetime import datetime
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logging.basicConfig(level=logging.DEBUG)


i = 0
l= open('Links/Com_Arenda.txt').read().splitlines()

page = l[i] 
oper = u'Аренда'





while i < len(l):
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Theproperty_Com(Spider):
	  def prepare(self):
	       self.f = page
	       for p in range(1,21):
		    try:
			 time.sleep(1)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 print g.doc.code
			 if g.doc.code ==200:
			      self.num = re.sub('[^\d]','',g.doc.select(u'//p[@class="quant"]').text().split(', ')[0])
			      self.pag = int(math.ceil(float(int(self.num))/float(15)))
			      self.dt = g.doc.select(u'//p[@class="current-city"]/a').text().replace(',','').replace(u' г','')
			      print self.dt,self.num
			      #link_sub = 'https://www.alta.ru/kladrs/search/?s_object_rf='+dt
			      link_sub = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+self.dt
			      g.go(link_sub) 
			      #b = g.response.json["results"][0]["formatted_address"]
			      #if int(len(b.split(', ')))==3:
				   #self.sub = b.split(', ')[1]
			      #elif int(len(b.split(', ')))==2:
				   #self.sub = b.split(', ')[0]
			      #self.sub = g.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["Address"]["Components"][2]["name"]    
			      self.sub = g.doc.rex_text(u'AdministrativeAreaName":"(.*?)"')
			      #try:
				   #self.sub = g.doc.select(u'//h1[@class="blue"]/following-sibling::div/ul/li[1]/span/a').text().split(', ')[0].replace(u' Город','')#reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(u' областьская ',' ').replace(u' областьая ',' ').replace(u' крайский ',' ')
			      #except IndexError:
				   #self.sub = g.doc.select(u'//div[@class="js_autoComplete"]/input/@value').text().split(', ')[0].replace(u' Город','')
				   
			      print self.sub,self.pag
			      del g
			      break
		    except(GrabTimeoutError,GrabNetworkError,IndexError,KeyError,AttributeError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
                    
	       else:
		    self.sub = ''
		    self.pag = 0	       
	       
	       self.workbook = xlsxwriter.Workbook(u'com/Theproperty_'+oper+str(i+1) + '.xlsx')
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
	       self.ws.write(0, 38, u"ОХРАНА")
	       self.ws.write(0, 39, u"КОНДИЦИОНИРОВАНИЕ")
	       self.ws.write(0, 40, u"ИНТЕРНЕТ")
	       self.ws.write(0, 41, u"ТЕЛЕФОН (КОЛИЧЕСТВО ЛИНИЙ)")
	       self.ws.write(0, 42, u"УСЛУГИ")
	       self.ws.write(0, 43, u"НАЛИЧИЕ ОТДЕЛКИ ПОМЕЩЕНИЙ")
	       
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):
	       if self.pag == 0:
		    self.stop()
	       else:
		    for x in range(1,self.pag+1):
                         link = self.f+'?page='+str(x)
                         yield Task ('post',url=link,refresh_cache=True,network_try_count=100)
	          
	       
            
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//p[@class="address"]/a'):
                    ur = grab.make_url_absolute(elem.attr('href'))  
                    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
               
     
	  def task_item(self, grab, task):

	       try:
		    ter= grab.doc.select(u'//dt[contains(text(),"Район")]/following-sibling::dd[1][contains(text(),", ")]').text()
	       except IndexError:
		    ter =''
	            
	       try:
		    orentir = grab.doc.select(u'//dt[contains(text(),"Метро, ориентир")]/following-sibling::dd[1]').text()
	       except IndexError:
		    orentir = ''              
		 
	       try:
		    metro = grab.doc.select(u'//fieldset[@class="fldst-object-address"]/legend').text().replace(u'Продажа ','').replace(u'Аренда ','').replace(u'универсального ',u'универсальные ').split(' в ')[0]
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//dt[contains(text(),"Тип строения")]/following-sibling::dd[1]').text()
		 #print rayon
	       except IndexError:
		    metro_min = ''
		    
	       try:
	            klass = grab.doc.select(u'//dt[contains(text(),"Класс")]/following-sibling::dd[1]').text()
	       except IndexError:
	            klass = ''		    
		   
	       try:
		    metro_tr = grab.doc.select(u'//dt[contains(text(),"Год постройки")]/following-sibling::dd[1]').text()
	       except IndexError:
		    metro_tr = ''
	       try:
		    pot = grab.doc.select(u'//dt[contains(text(),"Метро, ориентир")]/following-sibling::dd[1][contains(text(),"м. ")]').text()
	       except IndexError:
		    pot = ''
	       try:
		    metro1 = grab.doc.select(u'//dt[contains(text(),"Метро, ориентир")]/following-sibling::dd[1][contains(text(),"м. ")]').text().split(' — ')[0]
	       except IndexError:
		    metro1 = ''
	       try:
	            metro2 = grab.doc.select(u'//dt[contains(text(),"Метро, ориентир")]/following-sibling::dd[1][contains(text(),"м. ")]').text().split(' — ')[1]
	       except IndexError:
	            metro2 = ''			    
               try:
	            sost = grab.doc.select(u'//dt[contains(text(),"Особые условия")]/following-sibling::dd[1]').text()
	       except IndexError:
		    sost = ''
	       try:
	            bez = grab.doc.select(u'//h1').text().split(u' — ')[1]
	       except IndexError:
	            bez = ''
		   
	       try:
		    try:
		         price = grab.doc.select(u'//p[@id="priceMulti_-1_1"]/strong[1]').text()
		    except IndexError:
                         price = grab.doc.select(u'//div[@class="cleaner"]/preceding-sibling::p[@class="descr"][1]/strong[1]').text()
	       except IndexError:
		    price = ''

     
	       try:
		    plosh_ob = grab.doc.select(u'//dt[contains(text(),"Площадь")]/following-sibling::dd[1]').text()
		  #print rayon
	       except IndexError:
		    plosh_ob = ''
     
	       
		    
	       try:
		    et = grab.doc.select(u'//dt[contains(text(),"Этаж")]/following-sibling::dd[1]').text().split('/')[0].replace(u'не указан','')
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		    etagn = grab.doc.select(u'//dt[contains(text(),"Этаж")]/following-sibling::dd[1]').text().split('/')[1].replace(u'не указано','')
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''
		     
	       
		
		     
	       try:
		    try:
		         opis = grab.doc.select(u'//h2[contains(text(),"Дополнительная информация")]/following-sibling::p[2]').text()
		    except IndexError:
		         opis = grab.doc.select(u'//h2[contains(text(),"Дополнительная информация")]/following-sibling::p').text()
	       except IndexError:
		    opis = ''
		
	       try:
		    phone = re.sub('[^\d\+\,]','',grab.doc.select(u'//p[@class="phone"]').text())
	       except IndexError:
	            phone = ''
		   
	       try:
		    lico = grab.doc.select(u'//p[@class="name"]/a').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//p[@class="company"]/a').text()
		 #print rayon
	       except IndexError:
		    comp = ''
		    
	       try:
		    data = grab.doc.select(u'//h1').text()
	       except IndexError:
		    data = ''
		    
	       try:
	            lat = grab.doc.select(u'//div[@id="objMap"]').attr('data-coords').split(', ')[0].replace('[','')
	       except IndexError:
	            lat = ''
	       try:
	            lng = grab.doc.select(u'//div[@id="objMap"]').attr('data-coords').split(', ')[1]
	       except IndexError:
	            lng = ''
		    
	       try:
	            ohr = grab.doc.select(u'//dt[contains(text(),"Охрана")]/following-sibling::dd[1]').text()
	       except IndexError:
	            ohr = ''
               try:
		    park = grab.doc.select(u'//dt[contains(text(),"Парковка")]/following-sibling::dd[1]').text()
	       except IndexError:
	            park = ''
		    
	       try:
	            inet = grab.doc.select(u'//dt[contains(text(),"Интернет")]/following-sibling::dd[1]').text()
	       except IndexError:
	            inet = ''		    
		    
	       try:
	            linii = grab.doc.select(u'//dt[contains(text(),"Телефония")]/following-sibling::dd[1]').text()
	       except IndexError:
	            linii = ''
		    
	       try:
	            sosto = grab.doc.select(u'//dt[contains(text(),"Состояние здания")]/following-sibling::dd[1]').text()
	       except IndexError:
	            sosto = ''		    
		    
	       
		    
	       
		   
		   
	       
		   
	      
	   
	       projects = {'sub': self.sub,
		           'teritor': ter,
	                   'orentir':orentir.replace(pot,''),
		           'metro': metro,
	                   'metro1': metro1,
	                   'metro2': metro2,
	                   'shir':lat,
	                   'dol':lng,
	                   'naz': metro_min,
		           'god': metro_tr,
		           'cena': price,
		           'plosh_ob':plosh_ob,
	                   'etach': et,
		           'etashost': etagn.replace('-','|'),
	                   'klass':klass,
	                   'sostoyan':sost,
	                   'ochrana':bez,
		           'opis':opis,
		           'url':task.url,
		           'phone':phone,
		           'lico':lico,
		           'company':comp,                  
	                   'ohrana':ohr,
	                   'parkovka':park,
	                   'internet':inet,
	                   'line':linii,
	                   'sosto':sosto,
	                   'data':data}
	       
	       try:
		    ad= grab.doc.select(u'//h1').text().split(' — ')[1]
		    link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ad
		    yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    yield Task('adres',grab=grab,project=projects)	       
	     
	  def task_adres(self, grab, task):
	       try:
		    ray= grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["AddressDetails"]["Country"]["AdministrativeArea"]["SubAdministrativeArea"]["SubAdministrativeAreaName"]
	       except (ValueError,IndexError,TypeError,KeyError,AttributeError):
		    ray = ''
	       try:
		    if self.dt == u'Москва':
		         punkt = u'Москва'
		    elif self.dt == u'Санкт-Петербург':
		         punkt = u'Санкт-Петербург'
		    elif self.dt == u"Севастополь":
			 punkt= u"Севастополь"			 
		    else:
			 punkt= grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["AddressDetails"]["Country"]["AdministrativeArea"]["SubAdministrativeArea"]["Locality"]["LocalityName"]
	       except (ValueError,IndexError,TypeError,KeyError,AttributeError):
	            punkt = ''
	       try:
		    try:
			 try:
	                      uliza= grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["AddressDetails"]["Country"]["AdministrativeArea"]["SubAdministrativeArea"]["Locality"]["DependentLocality"]["Thoroughfare"]["ThoroughfareName"]
		         except (ValueError,IndexError,TypeError,KeyError,AttributeError):
			      uliza= grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["AddressDetails"]["Country"]["AdministrativeArea"]["SubAdministrativeArea"]["Locality"]["Thoroughfare"]["ThoroughfareName"]
		    except (ValueError,IndexError,TypeError,KeyError,AttributeError):
			 uliza= grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["AddressDetails"]["Country"]["AdministrativeArea"]["Locality"]["Thoroughfare"]["ThoroughfareName"]     
	       except (ValueError,IndexError,TypeError,KeyError,AttributeError):
	            uliza = ''
	       try:
	            try:
			 try:
			      dom= grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["AddressDetails"]["Country"]["AdministrativeArea"]["SubAdministrativeArea"]["Locality"]["DependentLocality"]["Thoroughfare"]["Premise"]["PremiseNumber"]
			 except (ValueError,IndexError,TypeError,KeyError,AttributeError):
			      dom= grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["AddressDetails"]["Country"]["AdministrativeArea"]["SubAdministrativeArea"]["Locality"]["Thoroughfare"]["Premise"]["PremiseNumber"]
		    except (ValueError,IndexError,TypeError,KeyError,AttributeError):
		         dom= grab.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["AddressDetails"]["Country"]["AdministrativeArea"]["Locality"]["Thoroughfare"]["Premise"]["PremiseNumber"]
	       except (ValueError,IndexError,TypeError,KeyError,AttributeError):
	            dom = ''		    
	       
	       project2 ={'rayon': ray,
	                   'punkt':punkt,
	                   'ulica':uliza,
	                   'dom':dom}	       
	       
	     
	       yield Task('write',project=task.project,proj=project2,grab=grab)
	     
	     
	     
	     
	     
	     
	  def task_write(self,grab,task):
	       
	       print('*'*50)	       
	       print  task.project['sub']
	       print  task.proj['rayon']
	       print  task.proj['punkt']
	       print  task.project['orentir']
	       print  task.project['teritor']
	       print  task.proj['ulica']
	       print  task.proj['dom']
	       print  task.project['metro']
	       print  task.project['naz']	      
	       print  task.project['god']
	       print  task.project['cena']	       
	       print  task.project['plosh_ob']	       
	       print  task.project['etach']
	       print  task.project['etashost']
	       print  task.project['klass']
	       print  task.project['sostoyan']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['ochrana']
	       print  task.project['shir']
	       print  task.project['dol']
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.proj['rayon'])
	       self.ws.write(self.result, 2,task.proj['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.proj['ulica'])
	       self.ws.write(self.result, 5,task.proj['dom'])
	       self.ws.write(self.result, 9,task.project['metro'])
	       self.ws.write(self.result, 8,task.project['naz'])
	       self.ws.write(self.result, 6,task.project['orentir'])
	       self.ws.write(self.result, 28,oper)
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 17, task.project['god'])
	       self.ws.write(self.result, 10, task.project['klass'])
	       self.ws.write(self.result, 14, task.project['plosh_ob'])
	       self.ws.write(self.result, 13, task.project['sostoyan'])
	       self.ws.write(self.result, 24, task.project['ochrana'])
	       self.ws.write(self.result, 26, task.project['metro1'])
	       self.ws.write(self.result, 27, task.project['metro2'])
	       self.ws.write(self.result, 15, task.project['etach'])
	       self.ws.write(self.result, 16, task.project['etashost'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'TheProperty.ru')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 33, task.project['data'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 34, task.project['shir'])
	       self.ws.write(self.result, 35, task.project['dol'])	       
	       self.ws.write(self.result, 37, task.project['parkovka'])
	       self.ws.write(self.result, 38, task.project['ohrana'])
	       self.ws.write(self.result, 40, task.project['internet'])
	       self.ws.write(self.result, 41, task.project['line'])
	       self.ws.write(self.result, 43, task.project['sosto'])
	      
	       
	       print('*'*50)
	       print self.sub
	      
	       print 'Ready - '+str(self.result)+'/'+self.num
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*50)
	       self.result+= 1
	       
	       
	       
	       
	       #if self.result > 10:
		    #self.stop()

     
     bot = Theproperty_Com(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     #command = 'mount -a'
     #os.system('echo %s|sudo -S %s' % ('1122', command))
     #time.sleep(2)
     bot.workbook.close()
     print('Done')
     #del bot
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
time.sleep(5)
os.system("/home/oleg/pars/property/comp.py")
     
     

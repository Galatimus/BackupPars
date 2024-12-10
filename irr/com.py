#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError,GrabTooManyRedirectsError
import logging
import base64
from grab import Grab
import time
import json
import math
import re
import os
from datetime import datetime
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


#g.proxylist.load_file(path='/home/oleg/Proxy/tipa.txt',proxy_type='http')

i = 0
l= open('Links/Comm.txt').read().splitlines()
dc = len(l)
page = l[i]



while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     
     class IRR_Com(Spider):
	  def prepare(self):
	       #self.count = 1 
	       self.f = page
	       if 'sale' in page:
		    self.oper = u'Продажа' 
	       else:
		    self.oper = u'Аренда' 
	       print self.oper
	       self.workbook = xlsxwriter.Workbook(u'Com/IRR_Коммерческая_'+str(i+1)+'.xlsx')
               self.ws = self.workbook.add_worksheet(u'Irr_Коммерческая')
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
	       
	       
	       self.conv = [(u' августа',u'.08.2019'), (u' июля',u'.07.2019'),
		    (u' мая',u'.05.2019'),(u' июня',u'.06.2019'),
		    (u' марта',u'.03.2019'),(u' апреля',u'.04.2019'),
		    (u' января',u'.01.2019'),(u' декабря',u'.12.2018'),
		    (u' сентября',u'.09.2019'),(u' ноября',u'.11.2018'),
		    (u' февраля',u'.02.2018'),(u' октября',u'.10.2018'), 
		    (u'сегодня,',datetime.today().strftime('%d.%m.%Y'))]
	       self.result= 1
	       
            
            
            
              
    
	  def task_generator(self):
	       #for x in range(1,self.pag+1):
                    #link = self.f+'page'+str(x)+'/'
                    #yield Task ('post',url=link.replace(u'page1/',''),refresh_cache=True,network_try_count=100)
               yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
	       
	       
	  def task_post(self,grab,task):
               if grab.doc.select(u'//h4[contains(text(),"Предложения рядом")]').exists()==True:
	            links = grab.doc.select(u'//h4[contains(text(),"Предложения рядом")]/preceding::a[contains(@class,"listing")]')
               else:
	            links = grab.doc.select(u'//a[@class="listing__itemTitle"]')
               for elem in links:
	            ur = grab.make_url_absolute(elem.attr('href'))
	            #print ur
	            yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
		    
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
                    ray =  grab.doc.select(u'//li[contains(text(),"АО:")]').text().split(': ')[1]
               except IndexError:
                    ray =''
               try:
                    punkt = grab.doc.rex_text(u'address_city":"(.*?)"').decode("unicode_escape")
               except (IndexError,TypeError,ValueError,KeyError):
                    punkt = ''
		    

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
                    godp = grab.doc.select(u'//li[contains(text(),"Год постройки/сдачи:")]').text().split(': ')[1]
               except IndexError:
                    godp = '' 
	       try:
                    naz = grab.doc.select(u'//li[contains(text(),"Назначение помещения:")]').text().split(': ')[1]
               except IndexError:
                    naz = ''  
		    
	       try:
	            vid = grab.doc.select(u'//li[contains(text(),"Тип здания:")]').text().split(': ')[1]
	       except IndexError:
	            vid = ''  		    
               try:
		    try:
                         price = grab.doc.select(u'//div[@class="productPage__price js-contentPrice"]').text()
		    except IndexError:
			 price = grab.doc.select(u'//div[@class="productPage__price"]').text()
               except IndexError:
                    price = ''
	    
	       try:
                    plosh = grab.doc.select(u'//li[contains(text(),"Общая площадь:")]').text().split(': ')[1]
               except IndexError:
                    plosh = ''
	    
	       try:
	            et = grab.doc.select(u'//li[contains(text(),"Этаж:")]').number()
               except IndexError:
                    et = ''
	       
	    
	       try:
                    et2 = grab.doc.select(u'//li[contains(text(),"Этажей в здании:")]').number()
               except IndexError:
                    et2 = ''
	    
	       try:
                    mat = grab.doc.select(u'//ul[@class="breadcrumbs__list js-breadcrumbsNav"]/li[4]/a/span').text()#.split(': ')[1]
               except IndexError:
                    mat = ''
	    
	       try:
                    pot = grab.doc.select(u'//li[contains(text(),"Комиссия:")]').text().split(': ')[1]
               except DataNotFound:
                    pot = ''
	    
	       try:
                    sost = grab.doc.select(u'//i[@class="icon icon_spot"]/following-sibling::div[1]').text()#.split(': ')[1]
               except DataNotFound:
                    sost = ''
		    
               try:
                    ohrana = grab.doc.select(u'//li[contains(text(),"Шоссе:")]').text().split(': ')[1]
               except DataNotFound:
                    ohrana =''
		    
               try:
                    gaz = grab.doc.select(u'//li[contains(text(),"Класс:")]').text().split(': ')[1]
               except DataNotFound:
                    gaz =''
		    
               try:
                    voda = grab.doc.select(u'//li[contains(text(),"Район города:")]').text().split(': ')[1]
               except DataNotFound:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//li[contains(text(),"До метро:")]').text().split(': ')[1]
               except DataNotFound:
                    kanal =''
		    
               try:
                    elekt = grab.doc.select(u'//h1').text()
               except DataNotFound:
                    elekt =''
		    
               try:
                    teplo = json.loads(grab.doc.select(u'//div[@class="js-productPageMap"]').attr('data-map-info'))['lat']
               except (IndexError,TypeError,ValueError,KeyError):
                    teplo =''
	       try:
	            lng = json.loads(grab.doc.select(u'//div[@class="js-productPageMap"]').attr('data-map-info'))['lng']
	       except (IndexError,TypeError,ValueError,KeyError):
		    lng = '' 
		    
               try:
                    opis = grab.doc.select(u'//h4[contains(text(),"Описание")]/following-sibling::p').text() 
               except DataNotFound:
                    opis = ''
		    
               try:
                    phone = re.sub('[^\d]', '',base64.b64decode(grab.doc.select('//input[@name="phoneBase64"]').attr('value')))
               except (AttributeError,DataNotFound):
                    phone = ''
	    
               try:
		    try:
                         lico = grab.doc.select(u'//input[@name="contactFace"]').attr('value')
		    except IndexError:
			 lico = grab.doc.select(u'//div[@class="productPage__infoTextBold productPage__infoTextBold_inline"]').text()
               except DataNotFound:
                    lico = ''
	    
               try:
                    com = grab.doc.select(u'//div[@class="productPage__infoTextBold productPage__infoTextBold_inline"]/a').text()
               except DataNotFound:
                    com = ''
               try:
		    data = datetime.strptime(grab.doc.rex_text(u'date_create":"(.*?)"}').split(' ')[0].replace('-','.'), '%Y.%m.%d')
               except DataNotFound:
                    data = ''
	       try:
		    d1 = grab.doc.select(u'//div[@class="productPage__createDate"]').text()
		    data1 = reduce(lambda d1, r: d1.replace(r[0], r[1]), self.conv, d1).replace(u'Размещено ','')
	       except IndexError:
	            data1 = '' 
		    
	       try:
	            park = grab.doc.select(u'//li[contains(text(),"Парковка")]').text().replace(u'Парковка',u'есть')
	       except DataNotFound:
	            park = ''		    
	    
               try:
                    metro = grab.doc.select(u'//li[contains(text(),"Метро:")]').text().split(': ')[1]
               except DataNotFound:
                    metro = ''

        
	
	
	
	       projects = {'url': task.url,
		           'phone': phone,
		           'price': price,
		           'opis': opis,
	                   'sub': sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'ulica': uliza,
		           'dom': dom,
	                    'naz': naz,
	                    'vid': vid,
		           'ploshad': plosh,
		           'et': et,
		           'ets': et2,
		           'mat': mat,
		           'potolok': pot,
		           'sost': sost,
		           'ochrana':ohrana,
		           'gaz': gaz,
		           'voda':voda,
		           'kanaliz':kanal,
		           'svet': elekt,
		           'teplo':teplo,
		           'god':godp,
		           'lico':lico.replace(com,''),
		           'komp':com,
		           'data':data.strftime('%d.%m.%Y'),
	                   'lng':lng,
	                   'metro':metro,
	                   'dataraz1': data1[:10],
		           'park':park}
	
	
	
               yield Task('write',project=projects,grab=grab)
	
  
	
	
	
	
	  def task_write(self,grab,task):
	      
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['god']
	       print  task.project['naz']
	       print  task.project['vid']
	       print  task.project['price']
	       print  task.project['ploshad']
	       print  task.project['et']
	       print  task.project['ets']
	       print  task.project['mat']
	       print  task.project['potolok']
	       print  task.project['sost']
	       print  task.project['ochrana']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['svet']
	       print  task.project['metro']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['opis']
	       print  task.project['lico']
	       print  task.project['komp']
	       print  task.project['data']
	       print  task.project['dataraz1']
	       print  task.project['park']
	       print  task.project['teplo']
	       print  task.project['lng']
	
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 3, task.project['rayon'])
	       self.ws.write(self.result, 26, task.project['metro'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	      
	       self.ws.write(self.result, 8, task.project['vid'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       self.ws.write(self.result, 35, task.project['lng'])
	       self.ws.write(self.result, 11, task.project['price'])
	       self.ws.write(self.result, 14, task.project['ploshad'])	
	       self.ws.write(self.result, 15, task.project['et'])
	       self.ws.write(self.result, 16, task.project['ets'])
	       self.ws.write(self.result, 17, task.project['god'])
	       self.ws.write(self.result, 7, task.project['mat'])
	       self.ws.write(self.result, 13, task.project['potolok'])
	       self.ws.write(self.result, 24, task.project['sost'])
	       self.ws.write(self.result, 36, task.project['ochrana'])
	       self.ws.write(self.result, 10, task.project['gaz'])
	       self.ws.write(self.result, 3, task.project['voda'])
	       self.ws.write(self.result, 27, task.project['kanaliz'])
	       self.ws.write(self.result, 33, task.project['svet'])
	       self.ws.write(self.result, 34, task.project['teplo'])
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['komp'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'Из рук в руки')
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 28, self.oper)
	       self.ws.write(self.result, 37, task.project['park'])
	       self.ws.write(self.result, 30, task.project['dataraz1'])
	       print('*'*50)
	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print '*',i+1,'/',dc,'*'
	       print self.oper
	       print('*'*50)
	       self.result+= 1
	       
	       #if self.result > 10:
		    #self.stop()	       	       
	


     bot = IRR_Com(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     command = 'mount -a'
     os.system('echo %s|sudo -S %s' % ('1122', command))
     time.sleep(2)
     bot.workbook.close()
     print('Done')
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
        
 
 
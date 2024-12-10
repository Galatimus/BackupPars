#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from sub import conv
import re
import time
import random
import xlsxwriter
#import requests
from datetime import datetime,timedelta
import sys
import os
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'comm/0001-0062_00_C_001-0057_BBPORT.xlsx')


class Farpost_Com(Spider):
     def prepare(self):
	  #self.f = page
	  self.ws = workbook.add_worksheet(u'Bbport')
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
	  #self.r = conv     
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(1,189):#52
	       yield Task ('post',url='https://russia.bbport.ru/commercial/prodaja/?vt=1&obj=realty&region=1&city=-1&sp=1&oper_type=119&page=%d'%x,refresh_cache=True,network_try_count=100)
	  for x1 in range(1,73):#9
	       yield Task ('post',url='https://russia.bbport.ru/commercial/arenda/?vt=1&obj=realty&city=-1&sp=1&oper_type=118&page=%d'%x1,refresh_cache=True,network_try_count=100)
	  
	  
			      
     def task_post(self,grab,task):    
	  for elem in grab.doc.select(u'//a[@class="object__title wa"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
     
	          
        
     def task_item(self, grab, task):
	  
	    
	 
	  try:
               ray= grab.doc.select(u'//div[@class="map mb_45"]/text()').text().replace(u'Россия, ','').replace(' -','')
	  except IndexError:
	       ray = '' 
	       
	  try:
	       dt = ray.split(', ')[0]
	       sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(u'крайский','').replace(u'областьская','').replace(u'областьая','')
	  except IndexError:
	       sub = u'Москва'		  
	       
	  try:
	       trassa = grab.doc.select(u'//div[contains(text(),"Тип строения")]/following::div[2]').text()
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//a[@class="categories__link"]').text()#.split(', ')[1]
	  except IndexError:
	       udal = ''
          try:
               seg = grab.doc.select(u'//div[contains(text(),"Класс строения")]/following::div[2]').text()#.split(', ')[1]
          except IndexError:
               seg = ''	       
	       
	  try:
	       try:
	            price = grab.doc.select(u'//div[contains(text(),"Стоимость:")]/following::div[@class="countBox"][1]').text()#.replace(u'a',u'р.')
	       except IndexError:
		    price = grab.doc.select(u'//div[contains(text(),"Стоимость аренды:")]/following::div[@class="countBox"][1]').text()
	  except IndexError:
	       price = ''
	       
	  try:
	       plosh = grab.doc.select(u'//div[contains(text(),"Площадь:")]/following::div[@class="countBox"][1]').text()
	  except IndexError:
	       plosh = '' 
	  try:
	       cena_za = grab.doc.select(u'//div[contains(text(),"Этаж")]/following::div[2]').text().split('/')[0]
	  except IndexError:
	       cena_za = '' 
	       
	  
	  try:
	       ohrana = grab.doc.select(u'//div[contains(text(),"Этаж")]/following::div[2]').text().split('/')[1]
	  except IndexError:
	       ohrana =''
	  try:
	       gaz = grab.doc.select(u'//div[contains(text(),"Год постройки")]/following::div[2]').text()
	  except IndexError:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//h2[contains(text(),"Контактные данные")]/following-sibling::a').text()
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//h2[contains(text(),"данные агентства")]/following-sibling::a').text()
	  except IndexError:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//div[@class="metro__text"]').text()
	  except IndexError:
	       elek =''
	  try:
	       teplo = grab.doc.select(u'//h1').text()
	  except IndexError:
	       teplo =''  
		    
	  try:
	       opis = grab.doc.select(u'//h2[contains(text(),"Дополнительная информация")]/following-sibling::div/div').text() 
	  except IndexError:
	       opis = ''
	       
	  #try:
	       #ph= grab.doc.select(u'//div[@class="phone togglePhone"]').attr('data-id')
	       #url_ph = 'https://russia.bbport.ru/clients/getphone/?JsHttpRequest='+ph+'-xml&id='+ph
               #r= requests.post(url_ph,verify=True,allow_redirects=False,timeout=15000)
	       #phone = re.sub('[^\d\+]','',r.content.split('phone')[1])
	       #del r
	  #except IndexError:
	       #phone = ''	       
	 	       
	  
	  try:
               if 'prodaja' in task.url:
	            oper = u'Продажа' 
               elif 'arenda' in task.url:
	            oper = u'Аренда'     
          except IndexError:
	       oper = ''
	       
	  try:
	       data = grab.doc.select(u'//div[@class="time"]').text().replace('-','.')
	    #print data
	  except IndexError:
	       data = ''
		    
	  
						   
	       
	  projects = {'url': task.url,
	              'sub': sub,
                      'rayon': ray,
                      'trassa': trassa,
                      'udal': udal,
	              'segment': seg,
                      'cena': price,
                      'plosh':plosh,
	              #'etah':ets,
	              #'god':god,
	              #'mat':mat,
	              'phone': random.choice(list(open('../phone.txt'))),
                      'cena_za': cena_za,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda.replace(kanal,''),
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'operazia':oper,
                      'data':data }
	  
	  
	  try:
	       #ad= 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ray.split(', ')[0]
	       link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ray
	       yield Task('sub',url=link,project=projects,refresh_cache=True,network_try_count=100)
               #yield Task('sub2',url=ad,refresh_cache=True,network_try_count=100)	       
	  except IndexError:
	       yield Task('sub',grab=grab,project=projects)
	       #yield Task('sub2',grab=grab)
	  
	  
	  
     #def task_sub2(self, grab, task):
	  #try:
	       #self.sub = grab.response.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["Address"]["Components"][2]["name"]
	  #except (IndexError,IndexError,KeyError,AttributeError):
	       #self.sub = ''
	  
 
     def task_sub(self, grab, task):
	
	  try:   
	       punkt= grab.doc.rex_text(u'LocalityName":"(.*?)"')
	  except IndexError:
	       punkt = ''
	  try:
	       ter=  grab.doc.rex_text(u'DependentLocalityName":"(.*?)"')
	  except IndexError:
	       ter =''
	  try:
	       uliza=grab.doc.rex_text(u'ThoroughfareName":"(.*?)"')
	  except IndexError:
	       uliza = ''
	  try:
	       dom = grab.doc.rex_text(u'PremiseNumber":"(.*?)"')
	  except IndexError:
	       dom = ''
     
	  project2 ={'punkt':punkt,
                     'teritor': ter,
                     'ulica':uliza,
                     'dom':dom}		       
	  
	  
	  yield Task('write',project=task.project,proj=project2,grab=grab)
	  
	       
 
            
     def task_write(self,grab,task):
	  print('*'*50)
	  print  task.project['sub']
	  print  task.project['rayon']
	  print  task.proj['punkt']
	  print  task.proj['teritor']
	  print  task.proj['ulica']
	  print  task.proj['dom']
	  print  task.project['trassa']
	  print  task.project['udal']
	  print  task.project['segment']
	  print  task.project['cena']
	  print  task.project['plosh']
	  #print  task.project['etah']
	  #print  task.project['god']
	  #print  task.project['mat']
	  #print  task.project['sostoyanie']
	  print  task.project['ohrana']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanaliz']
	  print  task.project['electr']
	  print  task.project['opis']
	  print task.project['url']
	  print  task.project['phone']
	  print  task.project['data']
          print  task.project['teplo']
	  
	  #global result
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 24, task.project['rayon'])
	  self.ws.write(self.result, 2, task.proj['punkt'])
	  self.ws.write(self.result, 3, task.proj['teritor'])
	  self.ws.write(self.result, 4, task.proj['ulica'])
	  self.ws.write(self.result, 10, task.project['segment'])
	  self.ws.write(self.result, 8, task.project['trassa'])
	  self.ws.write(self.result, 9, task.project['udal'])
	  self.ws.write(self.result, 5 , task.proj['dom'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  self.ws.write(self.result, 15, task.project['cena_za'])
	  self.ws.write(self.result, 16, task.project['ohrana'])
	  self.ws.write(self.result, 17, task.project['gaz'])
	  self.ws.write(self.result, 23, task.project['kanaliz'])
	  self.ws.write(self.result, 22, task.project['voda'])
	  #self.ws.write(self.result, 22, self.lico)
	  self.ws.write(self.result, 26, task.project['electr'])
	  self.ws.write(self.result, 33, task.project['teplo'])
	  self.ws.write(self.result, 19, u'BBport')
	  self.ws.write_string(self.result, 20, task.project['url'])
	  self.ws.write(self.result, 18, task.project['opis'])
	  self.ws.write(self.result, 21, task.project['phone'])
	  #self.ws.write(self.result, 25, self.sub+' ,'+task.project['teritor']+' ,'+task.project['teplo'])
	  self.ws.write(self.result, 29, task.project['data'])
	  self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 28, task.project['operazia'])
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  print  task.project['operazia']
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  #if self.result >50:
	       #self.stop()

     
bot = Farpost_Com(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
workbook.close()
print('Done!')
time.sleep(5)
os.system("/home/oleg/pars/small/doma24.py")





#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import os
import math
import random
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)






i = 0
l= open('Links/Comm.txt').read().splitlines()
page = l[i]





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class QP_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.workbook = xlsxwriter.Workbook(u'com/Yandex_Comm_'+str(i+1)+'.xlsx')
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
	       self.ws.write(0, 44, u"ОТДЕЛЬНЫЙ ВХОД")
	       self.ws.write(0, 45, u"ВЫСОТА ПОТОЛКОВ")
	       self.result= 1
	       
                
                
                
                
	  def task_generator(self):
	       for x in range(25):
		    yield Task ('post',url=self.f+'kupit/kommercheskaya-nedvizhimost/?page=%d'%x,refresh_cache=True,network_try_count=100)
	       for x1 in range(25):
		    yield Task ('post',url=self.f+'snyat/kommercheskaya-nedvizhimost/?page=%d'%x1,refresh_cache=True,network_try_count=100)
		    
               #yield Task ('post',url=self.f+'snyat/kommercheskaya-nedvizhimost/',refresh_cache=True,network_try_count=100)
	       #yield Task ('post',url=self.f+'kupit/kommercheskaya-nedvizhimost/',refresh_cache=True,network_try_count=100)
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//div[@class="OffersSerpItem__generalInfo"]/a'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	       #yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
		    
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//a[@class="Pager__radio-link"][contains(text(),"Следующая")]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print('*'*10)
		    print '!!','NO PAGE NEXT','!!'
		    print('*'*10)
		    
	  def task_item(self, grab, task):
	       try:
		    sub = grab.doc.select(u'//div[@class="OfferHeader"]/ol/li[1]/a/span').text()
	       except IndexError:
		    sub = ''	       	       
	       try:
	            mesto = grab.doc.select(u'//div[@class="OfferHeader__address"]').text()
	       except IndexError:
	            mesto =''
		    
	       try:
	            punkt = grab.doc.select(u'//div[@class="OfferHeader__address"]').text().split(', ')[0].replace(sub,'')
	       except IndexError:
	            punkt = ''	       
		
               try:
                    try:
                         ter =  grab.doc.select(u'//div[@class="OfferHeader"]/ol/li/a/span[contains(text(),"район")]').text()
                    except IndexError:
	                 ter =  grab.doc.select(u'//div[@class="OfferHeader"]/ol/li/a/span[contains(text(),"округ")]').text()
               except IndexError:
                    ter =''
               try:
                    uliza = grab.doc.select(u'//ul[@class="ColumnsList OfferTechDescription__list"]/li[2]/div/p').text().split(' из ')[0]
               except IndexError:
                    uliza = ''
               try:
                    dom = grab.doc.select(u'//div[@class="AuthorBadge__category"]').text()
               except IndexError:
                    dom = ''
		    
               try:
                    tip = grab.doc.select(u'//h1').text().split(', ')[0]
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//div[contains(text(),"Рекомендуемое назначение")]').text().split(': ')[1]
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//div[contains(text(),"Класс")]').text()
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//span[@class="price"]').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//ul[@class="ColumnsList OfferTechDescription__list"]/li[1]/div/p').text().split(' — ')[0]
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//ul[@class="ColumnsList OfferTechDescription__list"]/li[2]/div/p').text().split(' из ')[1]
               except IndexError:
                    ohrana =''
               try:
                    gaz =  grab.doc.select(u'//ul[@class="ColumnsList OfferTechDescription__list"]/li[4]/div/p').text()
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//span[@class="MetroStation__title"]').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//span[@class="MetroWithTime__distance OfferHeaderLocation__station-distance"]').text()
               except IndexError:
                    kanal =''
               try:
                    elek = grab.doc.rex_text(u'latitude":(.*?),')
               except IndexError:
                    elek =''
	       try:
	            lng = grab.doc.rex_text(u'longitude":(.*?),')
	       except IndexError:
	            lng =''		    
               try:
                    phone =  re.sub('[^\d\+]','',grab.doc.rex_text(u'phoneNumbers(.*?)redirectId'))
               except IndexError:
                    phone = random.choice(list(open('../phone.txt').read().splitlines()))

	       try:
		    opis = grab.doc.select(u'//p[@class="OfferTextDescription__text"]').text() 
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//div[@class="AuthorBadge__name"]').text()
	       except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//title').text().split(':')[0]
               except IndexError:
                    comp = ''
               try:
                    data1 = grab.doc.rex_text(u'updateDate":"(.*?)T').replace('-','.') 
               except IndexError:   
                    data1 = ''
	       try: 
	            data = grab.doc.rex_text(u'creationDate":"(.*?)T').replace('-','.')   
	       except IndexError:
		    data=''
		    
	       try:
	            oper = comp.split(' ')[0].replace(u'Снять',u'Аренда').replace(u'Купить',u'Продажа')
	       except IndexError:
	            oper = ''
		    
	       try:
	            mesto1 = grab.doc.select(u'//span[contains(text(),"Охраняемая парковка")]').text()
	       except IndexError:
	            mesto1 =''		    
	       try:
	            mesto2 = grab.doc.select(u'//span[contains(text(),"Охрана")]').text()
	       except IndexError:
		    mesto2 =''
	       try:
		    mesto3 = grab.doc.select(u'//span[contains(text(),"Кондиционер")]').text()
	       except IndexError:
	            mesto3 =''		    
	       try:
		    mesto4 = grab.doc.select(u'//ul[@class="ColumnsList OfferTechDescription__list"]/li[3]/div/p').text().split(' — ')[0]
	       except IndexError:
	            mesto4 =''     
	
               projects = {'sub': sub,
	                   'adress': mesto,
	                   'adress1': mesto1,
	                   'adress2': mesto2,
	                   'adress3': mesto3,
	                   'adress4': mesto4,
	                   'terit':ter, 
	                   'punkt': punkt.replace(ter,''),  
	                   'ulica':uliza,
	                   'dom':dom,
	                   'tip':tip,
	                   'naz':naz,
	                   'klass': klass,
	                   'cena': price,
	                   'plosh': plosh,
	                   'ohrana':ohrana,
	                   'gaz': gaz,
	                   'operacia': oper,
	                   'voda': voda,
	                   'kanaliz': kanal,
	                   'electr': elek,
	                   'opis': opis,
	                   'dol': lng,
	                   'url': task.url,
	                   'phone':phone[:12],
	                   'lico':lico,
	                   'company': comp,
	                   'data':data,
	                   'data1':data1}
	  
	  
	       yield Task('write',project=projects,grab=grab)
	  
	  
	  
	  def task_write(self,grab,task):
	       if task.project['sub'] <> '':
		    print('*'*10)	       
		    print  task.project['sub']
		    print  task.project['punkt']
		    print  task.project['adress']
		    print  task.project['terit']
		    print  task.project['ulica']
		    print  task.project['dom']
		    print  task.project['tip']
		    print  task.project['naz']
		    print  task.project['klass']
		    print  task.project['cena']
		    print  task.project['plosh']
		    print  task.project['gaz']
		    print  task.project['voda']
		    print  task.project['kanaliz']
		    print  task.project['electr']
		    print  task.project['ohrana']
		    print  task.project['opis']
		    print  task.project['url']
		    print  task.project['phone']
		    print  task.project['lico']
		    print  task.project['company']
		    print  task.project['data']
		    print  task.project['adress1']
		    
	       
		    
		    
	  
		    self.ws.write(self.result, 0, task.project['sub'])
		    self.ws.write(self.result, 24, task.project['adress'])
		    self.ws.write(self.result, 1, task.project['terit'])
		    self.ws.write(self.result, 2, task.project['punkt'])
		    self.ws.write(self.result, 15, task.project['ulica'])
		    self.ws.write(self.result, 23, task.project['dom'])
		    self.ws.write(self.result, 7, task.project['tip'])
		    self.ws.write(self.result, 9, task.project['naz'])
		    self.ws.write(self.result, 10, task.project['klass'])
		    self.ws.write(self.result, 11, task.project['cena'])
		    self.ws.write(self.result, 14, task.project['plosh'])
		    self.ws.write(self.result, 16, task.project['ohrana'])
		    self.ws.write(self.result, 17, task.project['gaz'])
		    self.ws.write(self.result, 26, task.project['voda'])
		    self.ws.write(self.result, 27, task.project['kanaliz'])
		    self.ws.write(self.result, 34, task.project['electr'])
		    self.ws.write(self.result, 18, task.project['opis'])
		    self.ws.write(self.result, 19, u'Яндекс Недвижимость')
		    self.ws.write_string(self.result, 20, task.project['url'])
		    self.ws.write(self.result, 21, task.project['phone'])
		    self.ws.write(self.result, 22, task.project['lico'])
		    self.ws.write(self.result, 33, task.project['company'])
		    self.ws.write(self.result, 29, task.project['data'])
		    self.ws.write(self.result, 30, task.project['data1'])
		    self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
		    self.ws.write(self.result, 28, task.project['operacia'])
		    self.ws.write(self.result, 35, task.project['dol'])
		    self.ws.write(self.result, 37, task.project['adress1'])
		    self.ws.write(self.result, 38, task.project['adress2'])
		    self.ws.write(self.result, 39, task.project['adress3'])
		    self.ws.write(self.result, 45, task.project['adress4'])
		    print('*'*10)
		    #print self.sub
		    print 'Ready - '+str(self.result)
		    print 'Tasks - %s' % self.task_queue.size() 
		    print '***',i+1,'/',len(l),'***'
		    print task.project['operacia']
		    print('*'*10)
		    self.result+= 1
     
		    
		    #if self.result > 10:
			 #self.stop()	       

     bot = QP_Com(thread_number=5, network_try_limit=1000)
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
     time.sleep(1)
     del bot
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
       
     
     
     
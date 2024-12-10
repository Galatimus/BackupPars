#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError
from grab import Grab
import logging
import base64
import random
import math
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
import os
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)









z = 0

s = ['ofisy','magaziny','proizvodstva','sklady','nezhilye-pomeshcheniya','torgovye-tsentry']

seg = s[z]


while True:
     print '****************************************',z+1,'/',len(s),'*******************************************'
     class Nedvizhka_Com(Spider):
	  def prepare(self):
	       self.url_do = 'https://tvoyadres.ru/nedvizhimost/'+seg+'/' 
	       for p in range(1,50):
		    try:
			 #time.sleep(1)
			 g = Grab()
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http') 
			 g.go(self.url_do)
			 print g.doc.code
			 self.num = re.sub('[^\d]', '',g.doc.select(u'//h2').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(10)))
			 print 'OK'
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,IndexError,KeyError,AttributeError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue     
	       else:
		    self.pag = 1
		    
	       if self.pag > 1000:
		    self.pag = 999
	       else:
		    self.pag = self.pag 
		    
	       print self.num,self.pag 	  
	       
	       self.workbook = xlsxwriter.Workbook(u'com/Tvoyadres_'+seg+ '_'+str(z+1)+'.xlsx')
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
	       self.conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
		       (u' мая ',u'.05.'),(u' июня ',u'.06.'),
		       (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
		       (u' января ',u'.01.'),(u' декабря ',u'.12.'),
		       (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
		       (u' февраля ',u'.02.'),(u' октября ',u'.10.'),
		       (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
		       (u'1.01.1970','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]	       
	       self.result= 1





	  def task_generator(self):
	       for x in range(self.pag):
		    yield Task ('post',url=self.url_do+'?page=%d'%x,refresh_cache=True,network_try_count=100)	       
	       
	      


	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//div[@class="title"]/a'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	       #yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)

	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//li[@class="next"]/a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print('*'*10)
		    print 'no_page'
		    print('*'*10)
		    logger.debug('%s taskq size' % self.task_queue.size())


	  def task_item(self, grab, task):
	       
	       try:
		    sub = grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]/a[contains(@href, "regiony")]').text()
	       except IndexError:
		    sub = ''	       
	       try:
		    mesto = grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]').text()
	       except IndexError:
	            mesto =''

	       try:
	            punkt = grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]/a[contains(@href, "goroda")]').text()
	       except IndexError:
	            punkt = ''

               try:
		    if grab.doc.select(u'//header[@class="property-title"]/figure/a[2][contains(text()," район")]').exists() == False:
			 ter =  grab.doc.select(u'//header[@class="property-title"]/figure/a[2]').text()
		    else:
			 ter =''
               except IndexError:
                    ter =''
               try:
                    uliza = grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]/a[contains(@href, "ulitsy")]').text()
               except IndexError:
                    uliza = ''
               try:
                    dom = grab.doc.select(u'//span[contains(text(),"Карта")]/following-sibling::span[1]/a[contains(@href, "doma")]').text()
               except IndexError:
                    dom = ''

               try:
                    tip = grab.doc.select(u'//span[contains(text(),"Объект")]/following-sibling::span').text()
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Этаж")]/following-sibling::dd[1]').text()
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Этажность")]/following-sibling::dd').text()
               except IndexError:
                    klass = ''
               try:
                    try:
                         price = re.sub('[^\d\.]', u'',grab.doc.select(u'//title').text().split(u'цен')[1].split(u',')[0]).replace('.',' руб.')
                    except IndexError:
	                 price = grab.doc.select(u'//meta[@name="description"]').attr('content').split(u'цене ')[1].split('.')[0]
               except IndexError:
                    price =''
               try:
                    plosh = grab.doc.select(u'//span[contains(text(),"Площадь объекта")]/following-sibling::span').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Год постройки")]/following-sibling::dd').text()
               except IndexError:
                    ohrana =''
               try:
                    gaz =  grab.doc.select(u'//span[contains(text(),"Сделка")]/following-sibling::span').text().replace(u'Продам',u'Продажа').replace(u'Сдам',u'Аренда')
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//span[@id="map"]').attr('data-coordinates').split(',')[0]
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//h1').text()
               except IndexError:
                    kanal =''
               elek =''
               try:
                    teplo = grab.doc.select(u'//span[@id="map"]').attr('data-coordinates').split(',')[1]
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//meta[@name="description"]').attr('content')
	       except IndexError:
	            opis = ''
               try:
		    try:
                         lico = grab.doc.rex_text(u'Собственник (.*?)"')
		    except IndexError:
			 lico = grab.doc.select(u'//a[contains(@href, "polzovateli")]').text()
               except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//dd[contains(text(),"Организация")]/following-sibling::dt[1]').text()
               except IndexError:
                    comp = ''
               try:
                    data1 = grab.doc.select(u'//span[contains(text(),"Дата обновления")]/following-sibling::span').text()
               except IndexError:
                    data1 = ''
	       try:
		    try:
			 try:
	                      data = grab.doc.select(u'//a[contains(@href, "uploads")]').attr('href').split('/uploads/')[1][:10].replace('/','.')
		         except IndexError:
			      data = grab.doc.select(u'//div[@class="image"]/ul/li/img[contains(@src, "jpg")]').attr('src').split('/uploads/')[1][:10].replace('/','.')
		    except IndexError:
			 data = grab.doc.select(u'//span[contains(text(),"Дата публикации")]/following-sibling::span').text()
	       except IndexError:
		    data=''

	       try:
                    phone = re.sub('[^\d]', '',base64.b64decode(grab.doc.rex_text(u'data-tel=(.*?)==')[1:]+'=='))
               except IndexError:
	            phone = random.choice(list(open('../phone.txt').read().splitlines()))
		    
	       oper = ''


	      
	       
	       #data1 = reduce(lambda data1, r: data1.replace(r[0], r[1]), self.conv, data1)
	       data = reduce(lambda data, r: data.replace(r[0], r[1]), self.conv, data)


               projects = {'sub': sub,
	                  'adress': mesto,
	                   'terit':ter,
	                   'punkt':punkt,
	                   'ulica':uliza,
	                   'dom':dom,
	                   'tip':tip,
	                   'naz':naz,
	                   'klass': klass,
	                   'cena': price,
	                   'plosh': plosh,
	                   'ohrana':ohrana,
	                   'gaz': gaz,
	                   'voda': voda,
	                   'operacia': oper,
	                   'kanaliz': kanal,
	                   'electr': elek,
	                   'teplo': teplo,
	                   'opis': opis,
	                   'url': task.url,
	                   'phone': phone,
	                   'lico':lico,
	                   'company': comp,
	                   'data':data.replace('25.10.2010',datetime.today().strftime('%d.%m.%Y')),
	                   'data1':data1}


	       yield Task('write',project=projects,grab=grab)



	  def task_write(self,grab,task):
	       #time.sleep(1)
	       print('*'*100)
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
	       self.ws.write(self.result, 24, task.project['adress'])
	       self.ws.write(self.result, 3, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 9, task.project['tip'])
	       self.ws.write(self.result, 15, task.project['naz'])
	       self.ws.write(self.result, 16, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 17, task.project['ohrana'])
	       self.ws.write(self.result, 28, task.project['gaz'])
	       self.ws.write(self.result, 34, task.project['voda'])
	       self.ws.write(self.result, 33, task.project['kanaliz'])
	       #self.ws.write(self.result, 34, task.project['electr'])
	       self.ws.write(self.result, 35, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'ТвойАдрес.RU')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 30, task.project['data1'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       #self.ws.write(self.result, 28, task.project['operacia'])
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+self.num
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',z+1,'/',len(s),'***'
	       print task.project['gaz']
	       print('*'*100)
	       self.result+= 1





	       #if self.result > 10:
	            #self.stop()

	       #if int(self.result) >= int(self.num)-1:
	            #self.stop()


     bot = Nedvizhka_Com(thread_number=5, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     try:
          bot.run()
     except KeyboardInterrupt:
          pass
     print('Save it...')
     time.sleep(2)
     bot.workbook.close()
     print('Done')
     z=z+1
     try:
	  seg = s[z]
     except IndexError:
	  break

	  






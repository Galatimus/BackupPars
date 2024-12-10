#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import math
from datetime import datetime,timedelta
import xlsxwriter
import random
import os
from cStringIO import StringIO
import pytesseract
from PIL import Image 
import time
import base64
from grab import Grab
import sys
reload(sys)
sys.setdefaultencoding('utf-8')



logging.basicConfig(level=logging.DEBUG)


i = 0
l= open('links/zem5.txt').read().splitlines()
dc = len(l)
page = l[i]

oper = u'Продажа'

while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class Kuz_zap(Spider):
     
     
	  def prepare(self):
	       self.f = page
	       self.link =l[i]	      
	       while True:
		    try:
			 time.sleep(2)
			 g = Grab(timeout=25, connect_timeout=65)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 self.sub = g.doc.rex_text(u'selected >(.*?)</option>')
			 self.num = re.sub('[^\d]','',g.doc.select(u'//span[@class="breadcrumbs-link-count js-breadcrumbs-link-count"]').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(50)))
			 print('*'*50)
			 print self.sub
			 print self.num
			 print self.pag
			 print('*'*50)
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue	        
	       self.workbook = xlsxwriter.Workbook(u'zem/Avito_%s' % bot.sub + u'_Земля_'+oper+'.xlsx')
	       self.ws = self.workbook.add_worksheet()
	       self.ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	       self.ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	       self.ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
	       self.ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	       self.ws.write(0, 4, u"УЛИЦА")
	       self.ws.write(0, 5, u"ДОМ")
	       self.ws.write(0, 6, u"ОРИЕНТИР")
	       self.ws.write(0, 7, u"ТРАССА")
	       self.ws.write(0, 8, u"УДАЛЕННОСТЬ")
	       self.ws.write(0, 9, u"ОПЕРАЦИЯ")
	       self.ws.write(0, 10, u"СТОИМОСТЬ")
	       self.ws.write(0, 11, u"ЦЕНА_ЗА_СОТКУ")
	       self.ws.write(0, 12, u"ПЛОЩАДЬ")
	       self.ws.write(0, 13, u"КАТЕГОРИЯ_ЗЕМЛИ")
	       self.ws.write(0, 14, u"ВИД_РАЗРЕШЕННОГО_ИСПОЛЬЗОВАНИЯ")
	       self.ws.write(0, 15, u"ГАЗОСНАБЖЕНИЕ")
	       self.ws.write(0, 16, u"ВОДОСНАБЖЕНИЕ")
	       self.ws.write(0, 17, u"КАНАЛИЗАЦИЯ")
	       self.ws.write(0, 18, u"ЭЛЕКТРОСНАБЖЕНИЕ")
	       self.ws.write(0, 19, u"ТЕПЛОСНАБЖЕНИЕ")
	       self.ws.write(0, 20, u"ОХРАНА")
	       self.ws.write(0, 21, u"ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	       self.ws.write(0, 22, u"ОПИСАНИЕ")
	       self.ws.write(0, 23, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 24, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 25, u"ТЕЛЕФОН")
	       self.ws.write(0, 26, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 27, u"КОМПАНИЯ")
	       self.ws.write(0, 28, u"ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 29, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 30, u"МЕСТОПОЛОЖЕНИЕ")
	       self.ws.write(0, 31, u"ШИРОТА")
	       self.ws.write(0, 32, u"ДОЛГОТА")
	       self.result= 1
	       
     
	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'?p=%d'%x,refresh_cache=True,network_try_count=100)
	       
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//a[@class="item-description-title-link"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur	      
		    yield Task('item',url=ur,refresh_cache=True,network_try_count=100)
	          
	  def task_item(self, grab, task):
	       try:
	            mesto =  grab.doc.select(u'//span[contains(text(),"Адрес")]/following-sibling::span').text()
	       except IndexError:
		    mesto = ''
	  
	       try:
		    ray =  grab.doc.select(u'//span[contains(text(),"Адрес")]/following-sibling::span').text().split(u'р-н ')[1].split(u', ')[0]
	       except IndexError:
		    ray = ''
	       try:
		    if self.sub == u'Москва':
			 punkt= u'Москва'
		    elif self.sub == u'Санкт-Петербург':
			 punkt= u'Санкт-Петербург'
		    elif self.sub == u'Севастополь':
			 punkt= u'Севастополь'
		    else:		    
			 punkt = grab.doc.rex_text(u'selected >(.*?)</option>')
	       except IndexError:
		    punkt = ''
	       try:
		    ul = grab.doc.select(u'//span[@itemprop="streetAddress"]').text()
		    uliza = re.split('(\W+)',ul)[1]
	       except IndexError:
		    uliza = ''
	       try:
		    d = grab.doc.select(u'//span[@itemprop="streetAddress"]').text()
		    dom =re.split('\W+', d,1)[1]
	       except IndexError:
		    dom = ''
	       try:
	            udal = grab.doc.select(u'//span[contains(text(),"Расстояние до города:")]/following-sibling::text()').text().replace(';','')
	       except IndexError:
	            udal = ''
	       try:
	            price = grab.doc.select('//span[@class="price-value-string js-price-value-string"]').text()
	       except IndexError:
	            price = ''
	       try:
	            price_sot = grab.doc.select(u'//li[@class="price-value-prices-list-item price-value-prices-list-item_size-small price-value-prices-list-item_pos-between"]').text().replace(u'за сотку ','')
	       except IndexError:
	            price_sot = ''
	       try:
	            plosh = grab.doc.select(u'//span[contains(text(),"Площадь:")]/following-sibling::text()').text()
	       except IndexError:
	            plosh = ''
	       try:
	            cat = grab.doc.select(u'//h1').text().split(' (')[1].replace(')','')
	       except IndexError:
	            cat = ''
	       try:
	            vid = grab.doc.select(u'//span[contains(text(),"Назначение участка:")]/following-sibling::text()').text()
	       except IndexError:
	            vid = ''
	       try:
	            opis = grab.doc.select(u'//div[@class="item-description"]/div').text() 
	       except IndexError:
	            opis = ''
	       try:
	            com = grab.doc.select(u'//div[contains(text(),"Агентство")]/preceding-sibling::div/descendant::a').text()
	       except IndexError:
	            com = ''

	       try:
		    try:
			 try:
			      lico = grab.doc.select(u'//div[contains(text(),"Частное лицо")]/preceding-sibling::div/descendant::a').text()
			 except IndexError:
			      lico = grab.doc.select(u'//div[contains(text(),"Арендодатель")]/following-sibling::div/div[1]').text()
	            except IndexError:
		         lico = grab.doc.select(u'//div[contains(text(),"Контактное лицо")]/following-sibling::div').text() 
	       except IndexError:
	            lico = ''
		    
	       try:
		    conv = [ (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
	                     (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1))),
	                     (u'июня', '.06.2019'),(u'июля', '.07.2019'),(u'августа', '.08.2019'),(u'января', '.01.2019'),(u'февраля', '.02.2019'),
		             (u'марта', '.03.2019'),(u'апреля', '.04.2019'),(u'мая', '.05.2019'),
	                     (u'ноября', '.11.2018'),(u'сентября', '.09.2019'),(u'октября', '.10.2018'),(u'декабря', '.12.2018')]
		    dt= grab.doc.select(u'//div[@class="title-info-metadata-item-redesign"]').text()#.split(u'размещено ')[1]
		    data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(' ','').split(u'в')[0]
	       except IndexError:
	            data = ''		    
	       try:
	            lat = grab.doc.rex_text(u'data-map-lat="(.*?)"')
	       except IndexError:
	            lat = ''
	       try:
		    lng = grab.doc.rex_text(u'data-map-lon="(.*?)"')
	       except IndexError:
		    lng = ''		    
	       
	       
	       
	       
	       projects = {'url': task.url,
	                 'sub': self.sub,
	                 'rayon': ray,
	                 'punkt': punkt,
	                 'ulica': uliza,
	                 'dom': dom,
	                 'udal': udal,
	                 'price': price,
	                 'price_sot': price_sot,
	                 'ploshad': plosh,
	                 'vid': vid,
	                 'cat': cat,
	                 'opis': opis,
	                 'lico':lico,
	                 'mesto':mesto,
	                 'company':com,
	                 'lat':lat,
	                 'lng':lng,	                 
	                 'dataraz': data }
	       
	       try:
		    #ad_id= re.sub(u'[^\d]','',task.url[-10:])
		    ad_id = re.sub(u'[^\d]','',grab.doc.rex_text(u'prodid(.*?)price'))
		    ad_phone = re.sub(u'[^0-9a-z]','',grab.doc.rex_text(u'avito.item.phone(.*?);'))
		    ad_subhash = re.findall(r'[0-9a-f]+', ad_phone)
		    if int(ad_id) % 2 == 0:
			 ad_subhash.reverse()
		    ad_subhash = ''.join(ad_subhash)[::3]
		    link = 'https://www.avito.ru/items/phone/'+ad_id+'?pkey='+ad_subhash
		    headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
			           'Accept-Encoding': 'gzip,deflate',
			           'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			           'Cookie': 'sessid='+ad_id+'.'+ad_subhash,
			           'Host': 'www.avito.ru',
			           'Referer': task.url,
			           'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0', 
			           'X-Requested-With' : 'XMLHttpRequest'}
		    gr = Grab()
		    gr.setup(url=link,headers=headers)
		    yield Task('phone',grab=gr,project=projects,refresh_cache=True,network_try_count=50)
	       except IndexError:
		    yield Task('phone',grab=grab,project=projects)
		    
	  def task_phone(self, grab, task):
	       try:
		    data_image64 = grab.doc.json['image64'].replace('data:image/png;base64,','') 
                    imgdata = base64.b64decode(data_image64)
                    im = Image.open(StringIO(imgdata))
                    x,y = im.size
                    phon = pytesseract.image_to_string(im.convert("RGB").resize((int(x*2), int(y*3)),Image.BICUBIC))
		    del im
	       except (IndexError,ValueError,GrabNetworkError,GrabTimeoutError,IOError):
	            phon = ''
		    
               phone=re.sub(u'[^\d]','',phon)
               if len(phone) > 11:
                    phone = random.choice(list(open('../phone.txt').read().splitlines()))

	       
	       yield Task('write',project=task.project,phone=phone,grab=grab)
	       
	       
	       
	       
	       
	  def task_write(self,grab,task):
	       if task.phone <> '':    
		    print('*'*50)
		    print  task.project['sub']
		    print  task.project['rayon']
		    print  task.project['punkt']
		    print  task.project['ulica']
		    print  task.project['dom']
		    print  task.project['udal']
		    print  task.project['price']
		    print  task.project['price_sot']
		    print  task.project['ploshad']
		    print  task.project['vid']
		    print  task.project['cat']
		    print  task.project['opis']
		    print  task.phone
		    print  task.project['lico']
		    print  task.project['company']
		    print  task.project['dataraz']
		    print  task.project['mesto']
		    print  task.project['url']
		    print  task.project['lat']
		    print  task.project['lng']
		   
		    self.ws.write(self.result, 0, task.project['sub'])
		    self.ws.write(self.result, 1, task.project['rayon'])
		    self.ws.write(self.result, 2, task.project['punkt'])
		    self.ws.write(self.result, 4, task.project['ulica'])
		    self.ws.write(self.result, 8, task.project['udal'])
		    self.ws.write(self.result, 9, oper)
		    self.ws.write(self.result, 10, task.project['price'])
		    self.ws.write(self.result, 11, task.project['price_sot'])
		    self.ws.write(self.result, 12, task.project['ploshad'])
		    self.ws.write(self.result, 14, task.project['vid'])
		    self.ws.write(self.result, 13, task.project['cat'])
		    self.ws.write(self.result, 22, task.project['opis'])
		    self.ws.write(self.result, 23, u'AVITO.RU')
		    self.ws.write_string(self.result, 24, task.project['url'])
		    self.ws.write(self.result, 25, task.phone)
		    self.ws.write(self.result, 26, task.project['lico'])
		    self.ws.write(self.result, 27, task.project['company'])
		    self.ws.write(self.result, 28, task.project['dataraz'])
		    self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
		    self.ws.write(self.result, 30, task.project['mesto'])
		    self.ws.write(self.result, 31, task.project['lat'])
		    self.ws.write(self.result, 32, task.project['lng'])
		  
		    
		    print('*'*50)
		    print 'Ready - '+str(self.result)+'/'+str(self.num)
		    print 'Tasks - %s' % self.task_queue.size()
		    print '***',i+1,'/',dc,'***'
		    print oper
		    print('*'*50)
		    self.result+= 1
	       
	       
	       
	            if str(self.result) == str(self.num):
		         self.stop()		  

     bot = Kuz_zap(thread_number=5, network_try_limit=1000)
     #bot.setup_queue(backend='mongo', database='Avitozem4',host='192.168.10.200')
     bot.load_proxylist('../tipa.txt','text_file',proxy_type='http')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     try:
          bot.run()
     except KeyboardInterrupt:
          pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     command = 'mount -a'# cifs //192.168.1.6/d /home/oleg/pars -o username=oleg,password=1122,iocharset=utf8,file_mode=0777,dir_mode=0777'
     #command = 'apt autoremove'
     p = os.system('echo %s|sudo -S %s' % ('1122', command))
     print p
     time.sleep(2)
     bot.workbook.close()
     #workbook.close()
     print('Done!')
     i=i+1
     try:
	  page = l[i]
     except IndexError:
	  break
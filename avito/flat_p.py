#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import math
import random
from datetime import datetime,timedelta
import xlsxwriter
from cStringIO import StringIO
import pytesseract
from PIL import Image 
import os
import time
import base64
from grab import Grab
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)



i = 0
l= open('links/kv_p.txt').read().splitlines()
dc = len(l)
page = l[i]  
oper = u'Продажа'
     


while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     

     class Dmir_Kv(Spider):
	  
	  
	  
          def prepare(self):
	       #self.count = 1 
	       self.f = page
	       self.link =l[i]
	       while True:
		    try:
			 time.sleep(2)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 self.sub = g.doc.rex_text(u'selected >(.*?)</option>')#.replace(u'Кызыл',u'Республика Тыва').replace(u'Долгопрудный',u'Московская область').replace(u'Гатчина',u'Ленинградская область')
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
	       self.workbook = xlsxwriter.Workbook(u'flats/Avito_%s' % bot.sub + u'_Жилье_'+oper+'.xlsx')
               self.ws = self.workbook.add_worksheet()
	       self.ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	       self.ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	       self.ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
	       self.ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	       self.ws.write(0, 4, u"УЛИЦА")
	       self.ws.write(0, 5, u"ДОМ")
	       self.ws.write(0, 6, u"ОРИЕНТИР")
	       self.ws.write(0, 7, u"СТАНЦИЯ_МЕТРО")
	       self.ws.write(0, 8, u"ДО_МЕТРО_МИНУТ")
	       self.ws.write(0, 9, u"ПЕШКОМ_ТРАНСПОРТОМ")
	       self.ws.write(0, 10, u"ТИП_ОБЪЕКТА")
	       self.ws.write(0, 11, u"ОПЕРАЦИЯ")
	       self.ws.write(0, 12, u"СТОИМОСТЬ")
	       self.ws.write(0, 13, u"ЦЕНА_М2")
	       self.ws.write(0, 14, u"КОЛИЧЕСТВО_КОМНАТ")
	       self.ws.write(0, 15, u"ПЛОЩАДЬ_ОБЩАЯ")
	       self.ws.write(0, 16, u"ПЛОЩАДЬ_ЖИЛАЯ")
	       self.ws.write(0, 17, u"ПЛОЩАДЬ_КУХНИ")
	       self.ws.write(0, 18, u"ПЛОЩАДЬ_КОМНАТ")
	       self.ws.write(0, 19, u"ЭТАЖ")
	       self.ws.write(0, 20, u"ЭТАЖНОСТЬ")
	       self.ws.write(0, 21, u"МАТЕРИАЛ_СТЕН")
	       self.ws.write(0, 22, u"ГОД_ПОСТРОЙКИ")
	       self.ws.write(0, 23, u"РАСПОЛОЖЕНИЕ_КОМНАТ")
	       self.ws.write(0, 24, u"БАЛКОН")
	       self.ws.write(0, 25, u"ЛОДЖИЯ")
	       self.ws.write(0, 26, u"САНУЗЕЛ")
	       self.ws.write(0, 27, u"ОКНА")
	       self.ws.write(0, 28, u"СОСТОЯНИЕ")
	       self.ws.write(0, 29, u"ВЫСОТА_ПОТОЛКОВ")
	       self.ws.write(0, 30, u"ЛИФТ")
	       self.ws.write(0, 31, u"РЫНОК")
	       self.ws.write(0, 32, u"КОНСЬЕРЖ")
	       self.ws.write(0, 33, u"ОПИСАНИЕ")
	       self.ws.write(0, 34, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 35, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 36, u"ТЕЛЕФОН")
	       self.ws.write(0, 37, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 38, u"КОМПАНИЯ")
	       self.ws.write(0, 39, u"ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 40, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 41, u"МЕСТОПОЛОЖЕНИЕ")
	      
	       self.result= 1
	      
            
            
            
              
    
	  def task_generator(self):
	       for x in range(1,self.pag+1):
	            yield Task ('post',url=self.f+'?p=%d'%x,refresh_cache=True,network_try_count=100)
        
        
            
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//a[@class="item-description-title-link"]'):
                    ur = grab.make_url_absolute(elem.attr('href'))  
                    yield Task('item',url=ur,refresh_cache=True,network_try_count=100)     
	         
	 
        
        
        
	  def task_item(self, grab, task):
	       try:
                    ray =  grab.doc.select(u'//span[@class="item-map-address"]/span[contains(text(), "р-н")]/text()').text().replace(',','')
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
                    uliza = grab.doc.select(u'//span[@itemprop="streetAddress"]').text()
	       except IndexError:
		    uliza = ''
               try:
                    dom =  re.sub('[^\d]','',grab.doc.select(u'//span[@class="item-phone-button-sub-text"]').text()+str(random.randint(1000000,9999999)))
               except IndexError:
                    dom = ''
		    
               try:
                    metro = grab.doc.select(u'//span[@class="item-map-metro"]').text().split(u' (')[0]
               except IndexError:
                    metro = ''
	       try:
                    metro_min = grab.doc.select(u'//span[@class="item-map-metro"]').text().split(' (')[1].replace(')','')
               except IndexError:
                    metro_min = ''
	       try:
                    metro_kak = grab.doc.select(u'//li[@class="metro"]/b[3]').text()
               except IndexError:
                    metro_kak = ''
               try:
                    tip_ob = u'Квартира'
               except IndexError:
                    tip_ob = ''
	       try:
	            price = grab.doc.select('//span[@class="price-value-string js-price-value-string"]').text()
               except IndexError:
	            price = ''
               try:
                    price_m = grab.doc.select('//li[@class="price-value-prices-list-item price-value-prices-list-item_size-small price-value-prices-list-item_pos-between"]').text()
               except IndexError:
                    price_m = ''
               try:
                    kol_komnat = grab.doc.select(u'//span[contains(text(),"Количество комнат:")]/following-sibling::text()').number()
               except IndexError:
                    kol_komnat = ''
               try:
                    et = grab.doc.select(u'//span[contains(text(),"Этаж:")]/following-sibling::text()').number()
               except IndexError:
                    et = ''
		    
               try:
                    et2 = grab.doc.select(u'//span[contains(text(),"Этажей в доме:")]/following-sibling::text()').number()
               except IndexError:
                    et2 = ''
		    
               try:
                    god = ''#grab.doc.rex_text(u'data-item-id="(.*?)"')[1:]
               except IndexError:
                    god = ''
		    
               try:
                    mat = grab.doc.select(u'//span[contains(text(),"Тип дома:")]/following-sibling::text()').text()
               except IndexError:
                    mat = ''
		    
               try:
                    pot = grab.doc.select(u'//li[contains(text(),"потолки")]/b').text()
               except IndexError:
                    pot = ''
		    
               try:
		    try:
                         sos = grab.doc.select(u'//a[@class="breadcrumb-link"][contains(text(), "Новостройки")]').text()
		    except IndexError:
			 sos = grab.doc.select(u'//a[@class="breadcrumb-link"][contains(text(), "Вторичка")]').text()
               except IndexError:
                    sos = ''
               try:
                    bal = grab.doc.select(u'//ul[@id="flat_data"]/li[contains(text(), "балкон")]/b').number()
               except IndexError:
                    bal = ''
               try:
                    logy = grab.doc.select(u'//ul[@id="flat_data"]/li[contains(text(), "лоджия")]/b').number()
               except IndexError:
                    logy = ''
               try:
                    plosh_ob = grab.doc.select(u'//span[contains(text(),"Общая площадь:")]/following-sibling::text()').text()
               except IndexError:
                    plosh_ob = ''
               try:
                    plosh_g = grab.doc.select(u'//span[contains(text(),"Жилая площадь:")]/following-sibling::text()').text()
               except IndexError:
                    plosh_g = ''
               try:
                    plosh_k = grab.doc.select(u'//span[contains(text(),"Площадь кухни:")]/following-sibling::text()').text()
               except IndexError:
                    plosh_k = ''
               try:
                    plosh_kom = grab.doc.select(u'//ul[@id="flat_data"]/li[contains(text(),"площадь комнат")]/b').text()
               except IndexError:
                    plosh_kom = ''               
               try:
                    san_u = grab.doc.select(u'//ul[@id="flat_data"]/li[contains(text(), "санузл")]/b').number()
               except IndexError:
                    san_u =''
               try:
                    okna = grab.doc.select(u'//ul[@id="flat_data"]/li[contains(text(), "окна")]/b').text()
               except IndexError:
                    okna =''
               try:
                    lift = grab.doc.select(u'//li[contains(text(),"лифт")]').number()
               except IndexError:
                    lift =''
               try:
                    kons = re.sub(u'^.*(?=консьерж)','', grab.doc.select(u'//*[contains(text(), "консьерж")]').text())[:8].replace(u'консьерж',u'есть')
               except IndexError:
	            kons = ''
               try:
	            opis = grab.doc.select(u'//div[@class="item-description"]/div').text() 
	       except IndexError:
	            opis = ''

               try:
                    try:
			 try:
			      lico = grab.doc.select(u'//div[contains(text(),"Продавец")]/following-sibling::div/div[1]').text()
			 except IndexError:
			      lico = grab.doc.select(u'//div[@class="seller-info-col"]/div[1]/div/a[contains(@href,"user")]').text()
		    except IndexError:
		         lico = grab.doc.select(u'//div[contains(text(),"Контактное лицо")]/following-sibling::div').text()
               except IndexError:
	            lico = ''
	    
	       try:
	            com = grab.doc.select(u'//div[@class="seller-info-name"]/a[contains(@href,"shopId")]').text()
               except IndexError:
	            com = ''
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
	            lin = []
		    for em in grab.doc.select(u'//div[@class="item-map-location"]/span[@itemprop="name"]'):
			 urr = em.text()
			 lin.append(urr)
		    data1 = ",".join(lin).replace(u'Адрес:,','')+','+grab.doc.select(u'//span[@class="item-map-address"]').text()
	       except IndexError:
	            data1 = ''
   
               try:
                    data2 =  grab.doc.select(u'//li[@id="history_wrap"]/table').text()
               except IndexError:
                    data2 = ''
	
	       projects = {'url': task.url,
		           'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'ulica': uliza,
	                   'dom': dom,
	                   'metro_min': metro_min,
	                   'metro': metro,
	                   'price': price,
	                   'price_m': price_m,
	                   'comnat': kol_komnat,
	                   'metro_kak': metro_kak,
	                   'object':tip_ob,
	                   'ploshad1': plosh_ob,
	                   'ploshad2': plosh_g,
	                   'ploshad3': plosh_k,
	                   'ploshad4': plosh_kom,
	                   'et': et,
	                   'ets': et2,
	                   'god': god,
	                   'balkon':bal,
	                   'logia':logy,
	                   'mat': mat,
	                   'potolok': pot,
	                   'sost': sos,
	                   'usel': san_u,
	                   'okna':okna,
	                   'lift': lift,
	                   'kons': kons,
	                   'opis': opis,
	                   'lico':lico,
	                   'company':com,
	                   'dataraz': data,
	                   'data1': data1,
	                   'data2': data2}
	       try:
		    #ad_id= re.sub(u'[^\d]','',task.url[-10:])
		    ad_id = re.sub(u'[^\d]','',grab.doc.rex_text(u'prodid(.*?)price'))
		    #ad_id= re.sub('[^\d]','',grab.doc.select(u'//div[@class="title-info-metadata-item"]').text().split(', ')[0])
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
			      'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0', 
			      'X-Requested-With' : 'XMLHttpRequest'}
		    gr = Grab()
		    gr.setup(url=link,headers=headers)
		    yield Task('phone',grab=gr,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
	            yield Task('phone',grab=grab,project=projects)	       
	
	
	  def task_phone(self, grab, task):
	       try:
		    data_image64 = grab.doc.json['image64'].replace('data:image/png;base64,','') 
		    imgdata = base64.b64decode(data_image64)
		    im = Image.open(StringIO(imgdata))
		    x,y = im.size
		    phon = pytesseract.image_to_string(im.convert("RGB").resize((int(x*2), int(y*3)),Image.BICUBIC))
	       except (IndexError,ValueError,GrabNetworkError,GrabTimeoutError,IOError):
		    phon = ''	  
	
	       phone=re.sub(u'[^\d]','',phon)
	       if phone == '05':
	            phone = task.project['dom']
		    
	       #elif phone == task.project['god']:
		    #phone = task.project['dom']
	
	
	       yield Task('write',project=task.project,phone=phone,grab=grab)
	

	
	
	
	
	  def task_write(self,grab,task):
	       if task.phone <> '':
		    print('*'*100)
		    print  task.project['sub']
		    print  task.project['rayon']
		    print  task.project['punkt']
		    print  task.project['ulica']
		    print  task.project['metro_min']
		    print  task.project['metro']
		    print  task.project['object']
		    print  task.project['price']
		    print  task.project['price_m']
		    print  task.project['comnat']
		    print  task.project['metro_kak']
		    print  task.project['ploshad1']
		    print  task.project['ploshad2']
		    print  task.project['ploshad3']
		    print  task.project['ploshad4']
		    print  task.project['et']
		    print  task.project['ets']
		    print  task.project['god']
		    print  task.project['balkon']
		    print  task.project['logia']
		    print  task.project['mat']
		    print  task.project['potolok']
		    print  task.project['sost']
		    print  task.project['usel']
		    print  task.project['okna']
		    print  task.project['lift']
		    print  task.project['kons']
		    print  task.project['opis']
		    print  task.project['url']
		    print  task.phone
		    print  task.project['lico']
		    print  task.project['company']
		    print  task.project['dataraz']
		    print  task.project['data1']
		    print  task.project['data2']
		    
		    self.ws.write(self.result,0, task.project['sub'])
		    self.ws.write(self.result,3, task.project['rayon'])
		    self.ws.write(self.result,2, task.project['punkt'])
		    self.ws.write(self.result,4, task.project['ulica'])
		    #self.ws.write(self.result,5, task.project['dom'])
		    self.ws.write(self.result,8, task.project['metro_min'])
		    self.ws.write(self.result,7, task.project['metro'])
		    self.ws.write(self.result,11, oper)
		    self.ws.write(self.result,12, task.project['price'])
		    self.ws.write(self.result,9, task.project['metro_kak'])
		    self.ws.write(self.result,10, task.project['object'])
		    #self.ws.write(self.result,12, task.project['ploshad1'])
		    self.ws.write(self.result,13, task.project['price_m'])
		    self.ws.write(self.result,14, task.project['comnat'])
		    self.ws.write(self.result,15, task.project['ploshad1'])
		    self.ws.write(self.result,16, task.project['ploshad2'])
		    self.ws.write(self.result,17, task.project['ploshad3'])
		    self.ws.write(self.result,18, task.project['ploshad4'])
		    self.ws.write(self.result,19, task.project['et'])
		    self.ws.write(self.result,20, task.project['ets'])
		    self.ws.write(self.result,21, task.project['mat'])
		    #self.ws.write(self.result,22, task.project['god'])
		    self.ws.write(self.result,24, task.project['balkon'])
		    self.ws.write(self.result,25, task.project['logia'])
		    self.ws.write(self.result,26, task.project['usel'])
		    self.ws.write(self.result,27, task.project['okna'])
		    self.ws.write(self.result,31, task.project['sost'])
		    self.ws.write(self.result,29, task.project['potolok'])
		    self.ws.write(self.result,30, task.project['lift'])
		    self.ws.write(self.result,32, task.project['kons'])
		    self.ws.write(self.result,33, task.project['opis'])
		    self.ws.write(self.result,34, u'AVITO.RU')
		    self.ws.write_string(self.result,35, task.project['url'])
		    self.ws.write(self.result,36, task.phone)
		    self.ws.write(self.result,37, task.project['lico'])
		    self.ws.write(self.result,38, task.project['company'])
		    self.ws.write(self.result,39, task.project['dataraz'])
		    self.ws.write(self.result,40, datetime.today().strftime('%d.%m.%Y'))
		    self.ws.write(self.result,41, task.project['data1'])
		    

		    print('*'*100)
		    print self.sub	       
		    print 'Ready - '+str(self.result)+'/'+str(self.num)
		    print 'Tasks - %s' % self.task_queue.size()
		    print '***',i+1,'/',dc,'***'
		    print oper
		    print('*'*100)
		    self.result+= 1
		    
		    #if self.result > 10:
			 #self.stop()
			 
		    if str(self.result) == str(self.num):
                         self.stop() 
	

     bot = Dmir_Kv(thread_number=5, network_try_limit=1000)
     #bot.setup_queue('mongo', database='AvitoFlat1',host='192.168.10.200')
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     p = os.system('echo %s|sudo -S %s' % ('1122', 'mount -a'))
     print p     
     time.sleep(2)     
     bot.workbook.close()
     print('Done!')     
     i=i+1
     try:
	  page = l[i]
     except IndexError:
          break

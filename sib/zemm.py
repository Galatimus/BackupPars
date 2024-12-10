#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
import os
import random
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


while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class Gdedom_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       for p in range(1,51):
                    try:
                         time.sleep(1)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http') 
                         g.go(self.f)
			 self.sub = g.doc.select(u'//div[@class="header-location"]/a').text()
			 print self.sub
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound, GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
		         continue
               else:
	            self.sub = ''

		    
	       
	              
	       self.workbook = xlsxwriter.Workbook(u'zem/Sibdom_Zemlya_'+str(i+1)+'.xlsx')
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
	       self.ws.write(0, 20, u"ПОДЪЕЗД")
	       self.ws.write(0, 21, u"ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	       self.ws.write(0, 22, u"ОПИСАНИЕ")
	       self.ws.write(0, 23, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 24, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 25, u"ТЕЛЕФОН")
	       self.ws.write(0, 26, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 27, u"КОМПАНИЯ")
	       self.ws.write(0, 28, u"ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 30, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
	       self.result= 1

    
	  def task_generator(self):
	       yield Task ('next',url=self.f,refresh_cache=True,network_try_count=100)
                  
        
        
        
          def task_next(self,grab,task):
	       for el in grab.doc.select(u'//select[@class="city-selector3-alter"]/option'):
		    urr = grab.make_url_absolute(el.attr('value'))  
		    #print urr
		    yield Task('goto', url=urr,refresh_cache=True,network_try_count=100)
	       
        
          def task_goto(self,grab,task):
	       for li in grab.doc.select(u'//div[@class="tab-content"]/div/ul/li/a[contains(@href,"zemlya")]'):
	            urlgo = grab.make_url_absolute(li.attr('href'))
	            #print urlgo
		    yield Task('post', url=urlgo,refresh_cache=True,network_try_count=100)
	       
            
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//div[@class="catalog-product-list catalog-product-list--lands"]/a'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task('page', grab=grab,refresh_cache=True,network_try_count=100)
	     
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//a[@class="page-pagination-next"]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print 'no_page'

	     
	  def task_item(self, grab, task):
	       
	       try:
		    ray = grab.doc.select(u'//div[@class="card-info-block"][1]/div[1]/div[2]').text().split(u', ')[1]
	       except IndexError:
	            ray = '' 
	       try:
		    punkt = grab.doc.select(u'//div[@class="card-info-block"][1]/div[1]/div[2]').text().split(u', ')[0]
	       except IndexError:
		    punkt = ''
		    
	       try:
	            ter= grab.doc.select(u'//span[contains(text(),"Ориентир")]/following::div[1]').text()
	       except IndexError:
		    ter =''
		    
	       try:
		    uliza = grab.doc.select(u'//span[contains(text(),"Адрес")]/following::div[1]').text().split(u', ')[0]
	       except IndexError:
	            uliza = ''
		    
	       try:
		    dom = grab.doc.select(u'//span[contains(text(),"Адрес")]/following::div[1]').text().split(u', ')[1]
	       except IndexError:
		    dom = ''
		    
	       try:
		    trassa = grab.doc.select(u'//div[@class="card-banner-price"]/following-sibling::p').text().split(u' за ')[0]
	       except IndexError:
		    trassa = ''
		    
	       try:
		    udal = grab.doc.select(u'//span[contains(text(),"Категория земель")]/following::div[1]').text()
	       except IndexError:
		    udal = ''
		    
	       try:
		    price = grab.doc.select(u'//div[@class="card-banner-price"]/span').text()
	       except IndexError:
		    price = ''
	       try:
		    plosh = grab.doc.select(u'//span[contains(text(),"Площадь")]/following::div[1]').text()
	       except IndexError:
		    plosh = ''

	       try:
		    vid = grab.doc.select(u'//span[contains(text(),"Водопровод")]/following::div[1]').text().replace(u'нет','')
	       except IndexError:
		    vid = '' 
		    
	       try:
		    oper = grab.doc.select(u'//ul[@class="breadcrumbs-list"]/li[2]/a').text().split(' ')[0]
	       except IndexError:
	            oper = ''	       
		    
	       try:
		    ohrana = grab.doc.select(u'//span[contains(text(),"Электричество")]/following::div[1]').text().replace(u'нет','')
	       except IndexError:
		    ohrana =''
	       try:
		    gaz = self.sub+', '+punkt+', '+ray+', '+uliza+' '+dom
	       except IndexError:
		    gaz =''
	       
	       try:
		    elek = grab.doc.select(u'//span[contains(text(),"Канализация")]/following::div[1]').text().replace(u'нет','')
	       except IndexError:
		    elek =''
		    
	       try:
		    lng = grab.doc.select(u'//span[contains(text(),"Подъезд")]/following::div[1]').text()
	       except IndexError:
	            lng =''		    
	       try:
		    teplo = grab.doc.select(u'//ul[@class="breadcrumbs-list"]/li[4]/a').text().split(u'добавлено ')[1]
	       except IndexError:
		    teplo =''
		    
	      
	       try:
		    opis = grab.doc.select(u'//div[@class="card-text card-description"]').text() 
	       except IndexError:
		    opis = ''
		    
	       try:
                    park = grab.doc.select(u'//ul[@class="breadcrumbs-list"]/li[4]/a').text().split(u'обновлено ')[1] 
	       except IndexError:
		    park = ''		    
		    
	       
	       data_id = re.sub('[^\d]','',task.url)
	       try:
		    data_key = grab.doc.select(u'//span[@class="js-show_phones"]').attr('data-key')
		    phone_url = 'https://www.sibdom.ru/api/showphone?id='+data_id+'&key='+data_key+'&owner=sticker'    
		    headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
			      'Accept-Encoding': 'gzip,deflate',
			      'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8=',
			      #'Cookie': 'sessid='+url1+'.'+url1,
			      'Host': 'www.sibdom.ru',
			      'Referer': task.url,
			      'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:52.0) Gecko/20100101 Firefox/52.0',
			      'X-Requested-With': 'XMLHttpRequest'}
		    g2 = grab.clone(headers=headers,proxy_auto_change=True)
		    g2.request(post=[('id',data_id), ('key', data_key),('owner', 'sticker')],headers=headers,url=phone_url)
		    phone =  re.sub('[^\d\+\,]','',re.findall('tel:(.*?)>',g2.doc.body)[0]) 
		    print 'Phone-OK'
		    del g2
	       except (IndexError,GrabConnectionError,GrabNetworkError,GrabTimeoutError):
	            phone = random.choice(list(open('../phone.txt').read().splitlines()))


     
	       try:
		    try:
		         lico = grab.doc.select(u'//div[contains(@class,"contacts")]').text()
		    except IndexError:
			 lico = grab.doc.select(u'//a[contains(@href,"agents/view")]').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//a[contains(@href,"organization/view")]').text()
	       except IndexError:
		    comp = ''
		    
	      
			 
	       
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
		           'gaz': gaz.replace(' , , ',''),
		           'electr': elek,
		           'teplo': teplo,
	                   'dol': lng,
	                   'opera': oper.replace(u'Сдача',u'Аренда'),
		           'opis':clearText,
		           'phone':phone,
	                   'parkov':park,
		           'lico':lico.replace(comp,''),
	                   'company':comp}
	       
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
	       print  task.project['electr']
	       print  task.project['dol']
	       print  task.project['teplo']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['gaz']
	       print  task.project['vid']
	       print  task.project['parkov']
	       
	       
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 3, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 6, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 11, task.project['trassa'])
	       self.ws.write(self.result, 13, task.project['udal'])
	       self.ws.write(self.result, 9, task.project['opera'])
	       self.ws.write(self.result, 10, task.project['cena'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 16, task.project['vid'])
	       self.ws.write(self.result, 31, task.project['gaz'])
	       self.ws.write(self.result, 17, task.project['electr'])
	       self.ws.write(self.result, 20, task.project['dol'])
	       self.ws.write(self.result, 29, task.project['parkov'])
	       self.ws.write(self.result, 28, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['ohrana'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'Сетевое издание «Сибдом»')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       #self.ws.write(self.result, 29, task.project['data'])
	       #self.ws.write(self.result, 30, task.project['data1'])
	       self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       #print task.sub
	       
	       print 'Ready - '+str(self.result)
	       print '*',i+1,'/',dc,'*'
	       print task.project['opera']
	       print('*'*50) 
	       self.result+= 1
		    

	       #if int(self.result) >= int(self.num)-3:
	            #self.stop()		    
     
	  
     bot = Gdedom_Zem(thread_number=5,network_try_limit=5000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
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
os.system("/home/oleg/pars/sib/comm.py")






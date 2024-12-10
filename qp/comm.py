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
	       self.link =l[i]
	      
               for p in range(1,51):
                    try:
                         time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http') 
                         g.go(self.f)
			 self.sub = g.doc.select(u'//a[@class="nav-link btn-region float-left jsc-region-selector"]/text()').text().replace(' / ','=')
                         self.num = re.sub('[^\d]','',g.doc.select(u'//strong[@class="items-count"]').text())
	                 self.pag = int(math.ceil(float(int(self.num))/float(30)))                         
                         print self.sub,self.pag,self.num
			 del g
                         break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.sub = ''
		    self.pag = 1
		    self.num = 1
		    
	       self.workbook = xlsxwriter.Workbook(u'com/Qp_%s' % bot.sub + u'_Коммерческая_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Qp_Коммерческая')
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
	       for x in range(self.pag+1):
                    yield Task ('post',url=self.f+'?offset='+str(x*30),refresh_cache=True,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@target="_self"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	      
            
	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[contains(text(),"район")]').text().split(' / ')[1]
	       except IndexError:
	            mesto =''
		    
	       try:
		    if self.sub == u"Москва":
			 punkt= u"Москва"
		    elif self.sub == u"Санкт-Петербург":
			 punkt='Санкт-Петербург'
	            else:
			 try:
	                      punkt = grab.doc.select(u'//span[contains(text(),"Населённый пункт")]/following::div[2]').text()
			 except IndexError:
			      punkt = grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[2]').text().split(' / ')[1]
	       except IndexError:
	            punkt = ''	       
		
               try:
                    ter =  grab.doc.select(u'//span[contains(text(),"Район города")]/following::div[2]').text()
               except IndexError:
                    ter =''
               try:
                    uliza = grab.doc.select(u'//span[contains(text(),"Улица")]/following::div[2]').text()
               except IndexError:
                    uliza = ''
               try:
                    dom = grab.doc.select(u'//span[contains(text(),"Номер дома")]/following::div[2]').text().replace('/','|')
               except IndexError:
                    dom = ''
		    
               try:
                    tip = grab.doc.select(u'//span[contains(text(),"Объект")]/following::div[2]').text()
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//span[contains(text(),"Назначение")]/following::div[2]').text()
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//div[@class="control-label"]/span[contains(text(),"таж")]/following::div[@class="controls"]').text()
               except IndexError:
                    klass = ''
               try:
		    try:
                         price = grab.doc.select(u'//a[@class="price a-preview-item"]').text().replace(' q',u' руб.')
		    except IndexError:
			 price = grab.doc.select(u'//div[@class="btn-group price-dropdown js-dropdown-openhover"]/button').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//span[contains(text(),"Площадь")]/following::div[2]').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//div[@class="control-label"]/span[contains(text(),"ность")]/following::div[2]').text()
               except IndexError:
                    ohrana =''
               try:
                    gaz =  u'Торг '+grab.doc.select(u'//span[contains(text(),"Торг")]/following::div[2]').text()
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//span[contains(text(),"Станция метро")]/following::div[2]/a').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//span[contains(text(),"Минут до метро")]/following::div[2]').text()
               except IndexError:
                    kanal =''
               try:
                    elek = re.sub('[^\d\.]','',grab.doc.select(u'//div[@id="ymap-details"]').attr('data-placemark').split(',')[0])
               except IndexError:
                    elek =''
	       try:
	            lng = re.sub('[^\d\.]','',grab.doc.select(u'//div[@id="ymap-details"]').attr('data-placemark').split(',')[1])
	       except IndexError:
	            lng =''		    
               try:
                    teplo =  grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[@class="controls"][1]').text().replace(' / ',', ')
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//span[contains(text(),"Дополнительная информация")]/following::div[2]').text() 
	       except IndexError:
	            opis = ''
               try:
		    try:
		         lico = grab.doc.select(u'//i[@class="fa fa-user"]/following::div[1]/div').text()
		    except IndexError:
			 lico = grab.doc.select(u'//div[@class="comment"]').text()
	       except IndexError:
                    lico = ''
               try:
                    co = grab.doc.select(u'//i[@class="fa fa-user"]/following::div[1]/div').text()
		    if "едвижимост" in co:
			 comp = co
		    else:
		         comp=''		    
               except IndexError:
                    comp = ''
               try:
                    data1 = grab.doc.select(u'//h1').text() 
               except IndexError:   
                    data1 = ''
	       try: 
	            data = grab.doc.select(u'//i[@class="fa fa-calendar "]/following-sibling::text()').text()
	       except IndexError:
		    data=''
	       
	       url1 = re.sub('[^\d]','',task.url)
	       try:
		    phone_url = 'https://qp.ru/viewadvert/ShowPhones?id='+url1+'&datatype=json'    
		    headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
			      'Accept-Encoding': 'gzip,deflate',
			      'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			      #'Cookie': 'QPSC4='+url1+'.'+url1,
			      'Host': 'qp.ru',
			      'Referer': task.url,
			      'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0',
			      'X-Requested-With': 'XMLHttpRequest'}
		    g2 = grab.clone(headers=headers,proxy_auto_change=True)
		    g2.request(headers=headers,url=phone_url) 
		    phone = ', '.join(g2.doc.json["phones"])
		    print 'Phone-OK'
		    del g2
	       except (IndexError,KeyError,ValueError,GrabConnectionError,GrabNetworkError,GrabTimeoutError):
		    del g2
		    phone = random.choice(list(open('../phone.txt').read().splitlines()))
          
	       try:
		    if 'prodau' in task.url:
		         oper = u'Продажа' 
	            elif 'sdau' in task.url:
		         oper = u'Аренда'
		    else:
		         oper = ''
	       except IndexError:
	            oper = ''	       
		    
     
	
               projects = {'sub': self.sub,
	                  'adress': mesto,
	                   'terit':ter, 
	                   'punkt':punkt.replace(mesto,''), 
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
	                   'teplo': teplo,
	                   'opis': opis,
	                   'dol': lng,
	                   'url': task.url,
	                   'phone': re.sub('[^\d\+\,]','',phone),
	                   'lico':lico.replace(comp,''),
	                   'company': comp,
	                   'data':data,
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
	       print  task.project['teplo']
	  
	       
	       
     
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['adress'])
	       self.ws.write(self.result, 3, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 8, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       self.ws.write(self.result, 15, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 16, task.project['ohrana'])
	       self.ws.write(self.result, 13, task.project['gaz'])
	       self.ws.write(self.result, 26, task.project['voda'])
	       self.ws.write(self.result, 27, task.project['kanaliz'])
	       self.ws.write(self.result, 34, task.project['electr'])
	       self.ws.write(self.result, 25, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'КУПИ.РУ')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 33, task.project['data1'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 28, task.project['operacia'])
	       self.ws.write(self.result, 35, task.project['dol'])
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+self.num
	       print 'Tasks - %s' % self.task_queue.size() 
	       print '***',i+1,'/',len(l),'***'
	       print task.project['operacia']
	       print('*'*100)
	       self.result+= 1

	       
	       #if self.result > 10:
	            #self.stop()	       

     bot = QP_Com(thread_number=10, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     command = 'mount -a'
     os.system('echo %s|sudo -S %s' % ('1122', command))
     time.sleep(2)
     bot.workbook.close()
     print('Done')
     del bot
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
       
     
     
     
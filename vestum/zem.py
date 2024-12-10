#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
from sub import conv
import re
import os
import math
import time
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






i = 0
l= open('links/zem.txt').read().splitlines()
dc = len(l)
page = l[i] 
oper = u'Продажа'


while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class Vestum_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       while True:
		    try:
			 time.sleep(2)
			 g = Grab(timeout=10, connect_timeout=10)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http')
			 g.go(self.f)
			 dt = g.doc.select(u'//span[@class="arrow"]/span').text()
			 self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(u' крайский ',' ')
			 self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="catalog-counter"]').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 print self.sub,self.pag,self.num
			 del g
			 break
	       
		    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       self.workbook = xlsxwriter.Workbook(u'zem/Vestum_%s' % bot.sub + u'_Земля_'+oper+str(i+1)+'.xlsx')
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
	       self.ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 30, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
		    
		    
	       self.result= 1
	     
		    
	 
	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'?page=%d'% x,refresh_cache=True,network_try_count=100)
	
	  def task_post(self,grab,task):
	 
	       for elem in grab.doc.select(u'//a[contains(@class,"card")]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	      
	     
	     
	  def task_item(self, grab, task):
	       try:
		    sub = grab.doc.select(u'//div[@id="content-address"]').text()
	       except IndexError:
		    sub = ''
	       try:
		    ray = grab.doc.select(u'//div[@id="content-address"]/a[contains(text(),"район")]').text()
		  #print ray 
	       except DataNotFound:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//div[@id="content-address"]').text().split(', ')[0]
	       except IndexError:
		    punkt = ''
		    
	       try:
		    ter= grab.doc.select(u'//div[@id="content-address"]').text().split(', ')[2]
	       except IndexError:
		    ter =''
		    
	       try:
		    
		    uliza = grab.doc.select(u'//div[@id="content-address"]').text().split(', ')[3]
		    #else:
			 #uliza = ''
	       except IndexError:
		    uliza = ''
		    
	       try:
		    dom = grab.doc.select(u'//div[@id="content-address"]').text().split(', ')[4]
	       except IndexError:
		    dom = ''
		    
	      
		    
	       try:
		    udal = grab.doc.select(u'//span[@class="price-metre"]').text()
	       except IndexError:
		    udal = ''
		    
	       try:
		    price = grab.doc.select(u'//span[@class="price-main"]').text()
	       except DataNotFound:
		    price = ''
	       try:
		    plosh = grab.doc.select(u'//td[contains(text(),"Площадь участка")]/following-sibling::td').text()
	       except DataNotFound:
		    plosh = ''
	       try:
		    vid = grab.doc.select(u'//td[contains(text(),"Целевое назначение")]/following-sibling::td').text()
	       except DataNotFound:
		    vid = '' 
		    
		    
	       try:
		    ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
	       except DataNotFound:
		    ohrana =''
	       try:
		    gaz = grab.doc.select(u'//td[contains(text(),"Газ")]/following-sibling::td').text()
	       except DataNotFound:
		    gaz =''
	       try:
		    voda = grab.doc.select(u'//td[contains(text(),"Водоснабжение")]/following-sibling::td').text()
	       except DataNotFound:
		    voda =''
	       try:
		    kanal = grab.doc.select(u'//td[contains(text(),"Канализация")]/following-sibling::td').text()
	       except DataNotFound:
		    kanal =''
	       try:
		    elek = grab.doc.select(u'//td[contains(text(),"Электричество")]/following-sibling::td').text()
	       except DataNotFound:
		    elek =''
	       try:
		    teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
	       except DataNotFound:
		    teplo =''
			 
	       try:
		    opis = grab.doc.select(u'//div[@class="text-info"]').text() 
	       except DataNotFound:
		    opis = ''
		    
	       try:
		    phone = re.sub('[^\d\,]','',grab.doc.select(u'//div[@id="seller-phone"]/span/div').attr('data-phone'))
	       except IndexError:
		    phone = ''
		    
	       try:
		    lico = grab.doc.select(u'//a[@class="seller-name"]').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//div[@class="seller-parent an"]').text()
	       except IndexError:
		    comp = ''
		    
	      
	       try:   
		    #dt= grab.doc.select(u'//td[contains(text(),"Обновлено")]/following-sibling::td').text()
		    data = grab.doc.select(u'//td[contains(text(),"Актуально")]/following-sibling::td/span').attr('data-datetime')[:10].replace(' ','.')
		 #print data
	       except IndexError:
		    data = ''
	       try:
	            #d1 = grab.doc.select(u'//td[contains(text(),"Размещено")]/following-sibling::td').text()#.split(u' в ')[0].replace(u'обновлено ','')
                    trassa = grab.doc.select(u'//td[contains(text(),"Размещено")]/following-sibling::td/span').attr('data-datetime')[:10].replace(' ','.')
	       except IndexError:
	            trassa = ''
			 
	       
							
		    
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
		           'gaz': gaz,
		           'voda': voda,
		           'kanaliz': kanal,
		           'electr': elek,
		           'teplo': teplo,
		           'opis':opis,
		           'phone':phone,
		           'lico':lico,
		           'company':comp,
		           'data':data,
		           'oper':sub
		           }
	       
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
	       print  task.project['vid']
	       print  task.project['ohrana']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['teplo']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['oper']
	       
	       #global result
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 28, task.project['trassa'])
	       self.ws.write(self.result, 11, task.project['udal'])
	       self.ws.write(self.result, 31, task.project['oper'])
	       self.ws.write(self.result, 10, task.project['cena'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 14, task.project['vid'])
	       self.ws.write(self.result, 15, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 18, task.project['electr'])
	       self.ws.write(self.result, 19, task.project['teplo'])
	       self.ws.write(self.result, 9, oper)
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'Вестум.RU')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       #print task.sub
	       
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '*',i+1,'/',dc,'*'
	       print oper
	       print('*'*50)	       
	       self.result+= 1
		    
		    
		    
	       #if self.result > 20:
		    #self.stop()
     
	  
     bot = Vestum_Zem(thread_number=5,network_try_limit=1000)
     #bot.setup_queue('mongo', database='Vestum',host='192.168.10.200')
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
     
time.sleep(5)
os.system("/home/oleg/pars/vestum/com.py")     







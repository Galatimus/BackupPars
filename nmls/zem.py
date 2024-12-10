#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError,GrabTooManyRedirectsError
import logging
from grab import Grab
import re
import os
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
l= open('Links/Zem.txt').read().splitlines()
dc = len(l)
page = l[i]
oper = u'Продажа'

while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class MK_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       for p in range(1,31):
		    try:
			 time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
                         self.sub = g.doc.select(u'//a[@class="dotted"]').text()
                         print self.sub
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabTooManyRedirectsError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.sub = ''
		    
	       self.workbook = xlsxwriter.Workbook(u'zem/Nmls_%s' % bot.sub + u'_Земля_'+oper+str(i)+'.xlsx')
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
	       yield Task ('post',url=page,refresh_cache=True,network_try_count=100)
		 
	  
	  
	  def task_post(self,grab,task):

	       for elem in grab.doc.select(u'//div[@class="font-weight-bold mb-1"]/a'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
		    
		    
	  def task_page(self,grab,task):
	       try:         
		    pg = grab.doc.select(u'//a[@class="nav-next"][contains(text(),"следующая")]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except DataNotFound:
		    print('*'*50)
		    print '!!!!!','NO PAGE NEXT','!!'
		    print('*'*50)
		    logger.debug('%s taskq size' % self.task_queue.size())             
	     
	     
		 
		 
	  
	     
	     
	  def task_item(self, grab, task):
	       
	       try:
		    ray = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/span/a[contains(@href,"raions")]').text().replace(u'район','')
		  #print ray 
	       except DataNotFound:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//td[contains(text(),"Город")]/following-sibling::td').text().replace(u'Область','')
	       except IndexError:
		    punkt = ''
		    
	       try:
		    try:
			 try:
		              ter = grab.doc.select(u'//td[contains(text(),"Район")]/following-sibling::td').text()
		         except DataNotFound:
		              ter = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/span/a[contains(text(),"поселок")]').text()
		    except DataNotFound:
			 ter = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/span/a[contains(text(),"деревня")]').text()
			 #ter= ''
	       except IndexError:
	            ter =''   
		    
	       try:
		    ul = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text()
		    if punkt <>'':
			 uliza = ul
		    else:
			 punkt = ul
			 uliza =''
			 
		    #try:
			 #try:
			      #uliza = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/span/a[contains(text(),"улица")]').text()
			 #except DataNotFound:
			      #uliza = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/span/a[contains(text(),"переулок")]').text()
		    #except DataNotFound:
			 #uliza = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/span/a[contains(text(),"проспект")]').text()
	       except DataNotFound:
	            uliza = '' 
		    
	       try:
		    dom = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text().split('д.')[1]
	       except IndexError:
		    dom = ''
		    
	       try:
		    trassa = grab.doc.select(u'//label[contains(text(),"Шоссе:")]/following-sibling::p').text().split(', ')[0]
		     #print rayon
	       except IndexError:
		    trassa = ''
		    
	       try:
		    udal = grab.doc.select(u'//label[contains(text(),"Шоссе:")]/following-sibling::p').text().split(', ')[1]
	       except IndexError:
		    udal = ''
		    
	       try:
		    price = grab.doc.select(u'//div[@class="card-price"]').text()
	       except DataNotFound:
		    price = ''
		    
	       
		    
	       try:
		    plosh = grab.doc.select(u'//td[contains(text(),"Площадь участка")]/following-sibling::td').text()
	       except DataNotFound:
		    plosh = ''
		    
	       
	       
	       
		    
	       try:
		    vid = grab.doc.select(u'//td[contains(text(),"Дата обновления")]/following-sibling::td').text().split(u' в ')[0]
	       except DataNotFound:
		    vid = '' 
		    
		    
	       try:
		    ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
	       except DataNotFound:
		    ohrana =''
	       try:
		    gaz = re.sub(u'^.*(?=газ)','', grab.doc.select(u'//*[contains(text(), "газ")]').text())[:3].replace(u'газ',u'есть')
	       except DataNotFound:
		    gaz =''
	       try:
		    voda = re.sub(u'^.*(?=вод)','', grab.doc.select(u'//*[contains(text(), "вод")]').text())[:3].replace(u'вод',u'есть')
	       except DataNotFound:
		    voda =''
	       try:
		    kanal = re.sub(u'^.*(?=санузел)','', grab.doc.select(u'//*[contains(text(), "санузел")]').text())[:7].replace(u'санузел',u'есть')
	       except DataNotFound:
		    kanal =''
	       try:
		    elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
	       except DataNotFound:
		    elek =''
	       try:
		    teplo =  grab.doc.select(u'//div[@id="YMapsID"]').attr('data-geocode')
	       except IndexError:
		    teplo =''
		    
	        
		   
			 
	       try:
		    opis = grab.doc.select(u'//div[@class="object-infoblock"][2]').text() 
	       except DataNotFound:
		    opis = ''
		    
	       try:
		    #try:
			 #phone = grab.doc.select(u'//b[contains(text(),"Контакт:")]/following-sibling::text()[2]').text().replace(' ','')
		    #except DataNotFound:
		    phone = re.sub('[^\d\,]','',grab.doc.select(u'//div[@class="mb5"]/img[contains(@src,"phone")]/following-sibling::text()').text())
	       except IndexError:
		    phone = ''
		    
	       try:
		    lico = re.sub('[\d]','',grab.doc.select(u'//div[contains(text(),"Агент:")]/a').text())
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//div[contains(text(),"АН:")]/a').text()
	       except IndexError:
		    comp = ''
		    
	       try:
		    
		    data= grab.doc.select(u'//td[contains(text(),"Дата подачи объявления")]/following-sibling::td').text().split(u' в ')[0]
	       except IndexError:
		    data = ''
			 
	       
							
		    
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
		           'data':data}
	                   
		           
	       
	       yield Task('write',project=projects,grab=grab)
		 
	  def task_write(self,grab,task):
	       if task.project['opis'] <> '':
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
		    print  task.project['gaz']
		    print  task.project['voda']
		    print  task.project['kanaliz']
		    print  task.project['electr']
		    print  task.project['opis']
		    print task.project['url']
		    print  task.project['phone']
		    print  task.project['lico']
		    print  task.project['company']
		    print  task.project['data']
		    print  task.project['vid']
		    print  task.project['teplo']
		    
		    
		    #global result
		    self.ws.write(self.result, 0, task.project['sub'])
		    self.ws.write(self.result, 1, task.project['rayon'])
		    self.ws.write(self.result, 2, task.project['punkt'])
		    self.ws.write(self.result, 3, task.project['teritor'])
		    self.ws.write(self.result, 4, task.project['ulica'])
		    self.ws.write(self.result, 5, task.project['dom'])
		    self.ws.write(self.result, 7, task.project['trassa'])
		    self.ws.write(self.result, 8, task.project['udal'])
		    self.ws.write(self.result, 9, oper)
		    self.ws.write(self.result, 10, task.project['cena'])
		    self.ws.write(self.result, 12, task.project['plosh'])
		    self.ws.write(self.result, 29, task.project['vid'])
		    self.ws.write(self.result, 15, task.project['gaz'])
		    self.ws.write(self.result, 16, task.project['voda'])
		    self.ws.write(self.result, 17, task.project['kanaliz'])
		    self.ws.write(self.result, 18, task.project['electr'])
		    self.ws.write(self.result, 31, task.project['teplo'])
		    self.ws.write(self.result, 20, task.project['ohrana'])
		    self.ws.write(self.result, 22, task.project['opis'])
		    self.ws.write(self.result, 23, u'Система "ИС Центр"')
		    self.ws.write_string(self.result, 24, task.project['url'])
		    self.ws.write(self.result, 25, task.project['phone'])
		    self.ws.write(self.result, 26, task.project['lico'])
		    self.ws.write(self.result, 27, task.project['company'])
		    self.ws.write(self.result, 28, task.project['data'])
		    self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
		    
		    print('*'*50)
		    print self.sub
		    print 'Ready - '+str(self.result)
		    logger.debug('Tasks - %s' % self.task_queue.size()) 
		    print '***',i+1,'/',dc,'***'
		    print oper
		    print('*'*50)
		    
		    self.result+= 1
			 
			 
			 
		    #if self.result > 20:
			 #self.stop()
     
	  
     bot = MK_Zem(thread_number=5,network_try_limit=1000)
     #bot.setup_queue(backend='mongo', database='nmls',host='192.168.10.200')
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     try:
	  command = 'mount -a'
	  os.system('echo %s|sudo -S %s' % ('1122', command))
	  time.sleep(2)
	  bot.workbook.close()
	  print('Done')
     except IOError:
	  time.sleep(30)
	  os.system('echo %s|sudo -S %s' % ('1122', 'mount -a'))
	  time.sleep(10)
	  bot.workbook.close()
	  print('Done!')
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break

time.sleep(5)
os.system("/home/oleg/pars/nmls/com.py")






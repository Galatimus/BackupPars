#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
import time
import xlsxwriter
import os
import math
import random
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






i = 0
l= open('links/Zagg.txt').read().splitlines()#.decode('cp1251').splitlines()
dc = len(l)
page = l[i]


while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class MK_Zag(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       for p in range(1,21):
		    try:
			 time.sleep(1)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 self.sub = self.f.split('/')[3].replace('+',' ')
		         #self.num = re.sub('[^\d]','',g.doc.select(u'//span[@class="count"]').text())
			 #self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 print self.sub#,self.num,self.pag
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
		         print g.config['proxy'],'Change proxy'
		         g.change_proxy()
		         continue
	       else:
	            self.sub = ''
	            self.pag = 0
	            self.num=0	
		    self.stop()
	       
	       self.workbook = xlsxwriter.Workbook(u'zagg/Mirkvartir_%s' % bot.sub + u'_Загород_'+str(i+1)+ '.xlsx')
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
	       self.ws.write(0, 9, u"КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
	       self.ws.write(0, 10, u"ТИП_ОБЪЕКТА")
	       self.ws.write(0, 11, u"ОПЕРАЦИЯ")
	       self.ws.write(0, 12, u"СТОИМОСТЬ")
	       self.ws.write(0, 13, u"ЦЕНА_М2")
	       self.ws.write(0, 14, u"ПЛОЩАДЬ_ОБЩАЯ")
	       self.ws.write(0, 15, u"КОЛИЧЕСТВО_КОМНАТ")
	       self.ws.write(0, 16, u"ЭТАЖНОСТЬ")
	       self.ws.write(0, 17, u"МАТЕРИАЛ_СТЕН")
	       self.ws.write(0, 18, u"ГОД_ПОСТРОЙКИ")
	       self.ws.write(0, 19, u"ПЛОЩАДЬ_УЧАСТКА")
	       self.ws.write(0, 20, u"ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	       self.ws.write(0, 21, u"ГАЗОСНАБЖЕНИЕ")
	       self.ws.write(0, 22, u"ВОДОСНАБЖЕНИЕ")
	       self.ws.write(0, 23, u"КАНАЛИЗАЦИЯ")
	       self.ws.write(0, 24, u"ЭЛЕКТРОСНАБЖЕНИЕ")
	       self.ws.write(0, 25, u"ТЕПЛОСНАБЖЕНИЕ")
	       self.ws.write(0, 26, u"ЛЕС")
	       self.ws.write(0, 27, u"ВОДОЕМ")
	       self.ws.write(0, 28, u"БЕЗОПАСНОСТЬ")
	       self.ws.write(0, 29, u"ОПИСАНИЕ")
	       self.ws.write(0, 30, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 31, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 32, u"ТЕЛЕФОН")
	       self.ws.write(0, 33, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 34, u"КОМПАНИЯ")
	       self.ws.write(0, 35, u"ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 36, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 37, u"МЕСТОПОЛОЖЕНИЕ")
	       
		    
		    
	       self.result= 1
	     
		    
	 
	  def task_generator(self):
	       for x in range(1,100):
                    yield Task ('post',url=self.f+'?p=%d'%x,network_try_count=100)
	
		 
	  def task_post(self,grab,task):
	       #if grab.doc.select(u'//h2[@class="nearby-header"]').exists()==True:
                    #links = grab.doc.select(u'//h2[@class="nearby-header"]/preceding::div[@class="item"]/a[1]')
               #else:
	            #links = grab.doc.select(u'//div[@class="item"]/a[1]')
               for elem in grab.doc.select(u'//a[@class="offer-title"]'):
	            ur = grab.make_url_absolute(elem.attr('href'))  
	            yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	     
	     
	  def task_item(self, grab, task):
	      
	       try:
		    ray = grab.doc.select(u'//p[@class="address"]/a[contains(text(),"р-н")]').text()
	       except IndexError:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[3]
	       except IndexError:
		    punkt = ''
		    
	       try:
		    ter= grab.doc.select(u'//div[@class="price m-m2"]').text()
	       except IndexError:
		    ter =''
		    
	       try:
                    uliza = grab.doc.select(u'//p[@class="address"]/a[contains(text(),"ул.")]').text()
	       except IndexError:
		    uliza = ''
		    
	       try:
		    dom = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[5].replace(u'цена','')
	       except IndexError:
		    dom = ''
		    
	       try:
		    trassa = grab.doc.select(u'//a[@class="m-highway"]').text()
	       except IndexError:
		    trassa = ''
		    
	       try:
		    udal = grab.doc.select(u'//label[contains(text(),"Шоссе:")]/following-sibling::p').text().split(', ')[1]
	       except IndexError:
		    udal = ''
	       try:
		    tip_ob = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[2]
	       except IndexError:
		    tip_ob = ''	       
		    
	       try:
		    price = grab.doc.select(u'//div[@class="price m-all"]').text()
	       except IndexError:
		    price = ''
	       try:
		    plosh = grab.doc.select(u'//span[contains(text(),"Общая площадь")]/preceding-sibling::strong').text()
	       except IndexError:
		    plosh = ''
		    
	       try:
		    kom = grab.doc.select(u'//p[@class="address"]').text()
	       except IndexError:
		    kom = ''	       
		    
	       try:
		    etash = grab.doc.select(u'//span[contains(text(),"этажа")]/preceding-sibling::strong').number()
	       except IndexError:
		    etash = ''
		    
	       try:
		    mat = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[1]
	       except IndexError:
		    mat = ''
	       try:
		    god = grab.doc.select(u'//label[contains(text(),"Дом:")]/following-sibling::p').text().split(', ')[1]
	       except IndexError:
		    god = ''	       
		    
	       try:
		    plosh_uch = grab.doc.select(u'//span[contains(text(),"площадь участка")]/preceding-sibling::strong').text()
	       except IndexError:
		    plosh_uch = ''
	       
		 
	       try:
		    ohrana = grab.doc.select(u'//span[contains(text(),"Безопасность")]/following-sibling::strong').text().replace(u'охрана',u'есть')
	       except IndexError:
		    ohrana =''
	       try:
		    z =  grab.doc.select(u'//span[contains(text(),"Коммуникации")]/following-sibling::strong').text()
		    if 'газ'in z:
			 gaz='есть'
		    else:
			 gaz =''
		    if 'вода' in z:
			 voda='есть'
		    else:
			 voda=''
                    if 'электричество' in z:
			 elek='есть'
		    else:
			 elek =''
                    if 'канализация' in z:
			 kanal='есть'
		    else:
			 kanal=''
		    if 'отопление' in z:
			 teplo ='есть'
		    else:
			 teplo ='' 
	       except (IndexError,UnboundLocalError):
		    gaz =''
		    voda=''
		    kanal=''
		    elek =''
		    teplo =''
		    z = ''
	       try:
		    les = re.sub(u'^.*(?=лес)','', grab.doc.select(u'//*[contains(text(), "лес")]').text())[:3].replace(u'лес',u'есть')
	       except IndexError:
		    les =''
		 
	       try:
		    vodoem = re.sub(u'^.*(?=озер)','', grab.doc.select(u'//*[contains(text(), "озер")]').text())[:4].replace(u'озер',u'есть')
	       except IndexError:
		    vodoem =''	  
	       try:
		    opis = grab.doc.select(u'//div[@class="l-object-description"]/p').text() 
	       except IndexError:
		    opis = ''

	       try:
		    lico = grab.doc.select(u'//div[@class="seller-info"]/p/strong').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//span[contains(text(),"Агентство недвижимости")]/preceding-sibling::strong').text()
	       except IndexError:
		    comp = ''
		    
	       try:
		    try:
                         data = grab.doc.select(u'//div[@class="l-object-dates"]/p[2]').text().split(u' размещено ')[1].split(u' в ')[0].replace(u'сегодня', (datetime.today().strftime('%d.%m.%Y')))
                    except IndexError:
	                 data = grab.doc.select(u'//div[@class="dates"]').text().split(u' размещено ')[1].split(u' в ')[0].replace(u'сегодня', (datetime.today().strftime('%d.%m.%Y')))
	       except IndexError:
	            data = ''
		    
	       
	       try:
                    phone = random.choice(list(open('../phone.txt').read().splitlines()))
	       except IndexError:
		    phone = ''    
			 
	       

		    
	       projects = {'url': task.url,
		           'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza.replace(punkt,''),
		            'dom': dom.replace(uliza,''),
		           'trassa': trassa,
		           'udal': udal,
		           'object': tip_ob,
		           'cena': price,
		           'plosh':plosh,
		           'kom':kom,
		           'etach': etash,
		           'material': mat,
		           'god_postr': god,
		           'plouh': plosh_uch,
		           'ohrana':ohrana,
		           'gaz': gaz,
		           'voda': voda,
		           'kanaliz': kanal,
		           'electr': elek,
		           'teplo': teplo,
	                   'phone':phone,
		           'les': les,
		           'vodoem':vodoem,	              
		           'opis':opis,
		           'lico':lico.replace(comp,''),
		           'company':comp,
		           'data':data }
	       
	      	       
	       
	 
	       
	       yield Task('write',project=projects,grab=grab)
		 
	  def task_write(self,grab,task):
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']       
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['trassa']
	       print  task.project['udal']
	       print  task.project['object']
	       print  task.project['cena']
	       print  task.project['teritor']
	       print  task.project['plosh']
	       print  task.project['kom']
	       print  task.project['etach']
	       print  task.project['god_postr']
	       print  task.project['plouh']
	       print  task.project['ohrana']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['teplo']
	       print  task.project['les']
	       print  task.project['vodoem']	  
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       
	       
	       #global result
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 13, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 7, task.project['trassa'])
	       self.ws.write(self.result, 8, task.project['udal'])
	       #self.ws.write(self.result, 11, oper)
	       self.ws.write(self.result, 10, task.project['object'])
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 37, task.project['kom'])
	       self.ws.write(self.result, 21, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['etach'])
	       self.ws.write(self.result, 11, task.project['material'])
	       self.ws.write(self.result, 18, task.project['god_postr'])
	       self.ws.write(self.result, 23, task.project['kanaliz'])
	       self.ws.write(self.result, 24, task.project['electr'])
	       self.ws.write(self.result, 19, task.project['plouh'])
	       self.ws.write(self.result, 28, task.project['ohrana'])
	       self.ws.write(self.result, 22, task.project['voda'])	  
	       self.ws.write(self.result, 25, task.project['teplo'])
	       self.ws.write(self.result, 26, task.project['les'])
	       self.ws.write(self.result, 27, task.project['vodoem'])
	       self.ws.write(self.result, 29, task.project['opis'])
	       self.ws.write(self.result, 30, u'MIRKVARTIR.RU')
	       self.ws.write_string(self.result, 31, task.project['url'])
	       self.ws.write(self.result, 32, task.project['phone'])
	       self.ws.write(self.result, 33, task.project['lico'])
	       self.ws.write(self.result, 34, task.project['company'])
	       self.ws.write(self.result, 35, task.project['data'])
	       self.ws.write(self.result, 36, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       #print task.sub
	       
	       print 'Ready - '+str(self.result)#+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '*',i+1,'/',dc,'*'
	       print task.project['material']
	       print('*'*50)	       
	       self.result+= 1
	       
	       #if self.result > 5000:
		    #self.stop()
               #if str(self.result) == str(self.num):
	            #self.stop()		    
     
	  
     bot = MK_Zag(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     try:
          bot.run()
     except KeyboardInterrupt:
          pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     bot.workbook.close()
     print('Done')
     i=i+1
     try:
	  page = l[i]
     except IndexError:
          break    
       







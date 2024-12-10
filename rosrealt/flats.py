#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
import logging
from datetime import datetime,timedelta
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import time
from grab import Grab
import re
import xlsxwriter
import sys
import os
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)

i = 0
l= open('links/Kvart.txt').read().splitlines()
page = l[i]


while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Rosreal_Kv(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       if 'prodam' in self.f:
		    self.oper = u'Продажа' 
	       elif 'arenda' in self.f:
	            self.oper = u'Аренда'
	       for p in range(1,51):
		    try:
			 time.sleep(1)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http') 
			 g.go(self.f)
			 self.sub = g.doc.select(u'//a[@class="a_cityp1"]').text()
			 print self.sub,' | ',self.oper
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
	            self.sub = ''
		    
	       self.workbook = xlsxwriter.Workbook(u'flats/Rosrealt_%s' % bot.sub + u'_Жилье_'+str(i+1)+'.xlsx')
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
	       self.ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
	       self.ws.write(0, 33, u"ОПИСАНИЕ")
	       self.ws.write(0, 34, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 35, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 36, u"ТЕЛЕФОН")
	       self.ws.write(0, 37, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 38, u"КОМПАНИЯ")
	       self.ws.write(0, 39, u"ДАТА_РАЗМЕЩЕНИЯ_ОБЪЯВЛЕНИЯ")
	       self.ws.write(0, 40, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 41, u"ДАТА_ПАРСИНГА")
	       self.result= 1
     
	  def task_generator(self):
	       yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
     
     
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//a[@class="nounder"][contains(@href, "kvartira")]'):
	            ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
	            if u'rosrealt' in ur:
		         yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	       
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//div[@class="nolink"][contains(text(),"Следующая страница")]/ancestor::a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print('*'*10)
		    print 'No_Page'
		    print('*'*10) 
	
	
	  def task_item(self, grab, task):
	       
	       try:
		    ray = grab.doc.select(u'//td/p[contains(text(),"площадь кухни")]').text().split(': ')[1] 
	       except IndexError:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//div[@id="path_block"]/div[2]/a/span').text()
	       except IndexError:
		    punkt = ''
		    
	       try:
		    ter = grab.doc.select(u'//p[@class="pbig_gray"]/b/a[contains(text(),"район")]').text()#.split(', ')[1]
               except IndexError:
		    ter =''
		    
	       try:
		    try:
		         uliza = grab.doc.select(u'//a[@class="nounder"][contains(@href, "?ul=")]').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//div[@id="colr"]/p[1]/a/following-sibling::text()').text().split(', ')[1]
	       except IndexError:
		    uliza = ''
	       
	       try:
		    dom = re.sub('[^\d\.]', u'',grab.doc.select(u'//p[contains(text(),"обновлено:")]').text())
	       except IndexError:
		    dom = ''
	       try:
		    tip_ob = u'Квартира'
 	       except IndexError:
		    tip_ob = ''
		    
	       try:
		    price = grab.doc.select(u'//p[@class="pbig_gray"]/b[contains(text(),"руб.")]/parent::p').text().split(' | ')[0]
	       except IndexError:
		    price = ''
		    
	       try:
		    cena_za = grab.doc.select(u'//td/p[contains(text(),"жилая площадь")]').text().split(': ')[1] 
	       except IndexError:
		    cena_za = ''	       
		    
	       try:
		    komnat = grab.doc.select(u'//td/p[contains(text(),"комнатная")]').number()
	       except IndexError:
		    komnat = ''
		    
	       try:
		    plosh_ob = grab.doc.select(u'//td/p[contains(text(),"общая площадь")]').text().split(': ')[1]
	       except IndexError:
		    plosh_ob = ''
		    
	       try:
		    etash = grab.doc.select(u'//td/p[contains(text(),"этаж:")]').number()
	       except IndexError:
		    etash = ''
		    
	       try:
	            ets = grab.doc.select(u'//nobr[contains(text(),"этажность:")]').number()
	       except IndexError:
	            ets = '' 
		    
	       try:
		    sost = grab.doc.select(u'//td/p[contains(text(),"ремон")]').text()
	       except IndexError:
		    sost = '' 
		    
	       try:
	            rinok = grab.doc.select(u'//td/p[contains(text(),"категория")]').text().split(': ')[1]
	       except IndexError:
		    rinok = ''
			 
	       try:
		    opis = grab.doc.select(u'//div[@class="info_self"]').text()   
	       except IndexError:
		    opis = ''
		    
	       try:
		    phone = grab.doc.select(u'//p[@class="pbig_gray_contact"]/a[contains(@href, "tel:")]').text()
	       except IndexError:
		    phone = ''
		    
	       try:
		    lico = grab.doc.select(u'//div[@class="kontakt"]/p[1]').text().split(', ')[0]
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//div[@class="kontakt"]/p[1]').text().split(', ')[1]
	       except IndexError:
		    comp = ''
		    
	       try:
		    data = re.sub('[^\d\.]', u'',grab.doc.select(u'//p[contains(text(),"добавлено:")]').text())
	       except IndexError:
		    data = ''
			 
	       try:
                    oper = grab.doc.select(u'//p[@class="pbig_gray"]/b[@class="blue"]').text()
	       except IndexError:
		    oper = ''		  
		    
		    
	       clearText = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", lico)
	       clearText = re.sub(u"[.,\-\s]{3,}", " ", clearText)      
		
	       projects = {'url': task.url,
		           'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt.replace(self.sub,''),
		           'teritor': ter,
		           'ulica': uliza.replace(ter,''),
		           'dom': dom,
		           'object': tip_ob,
		           'cena': price,
		           'cena_za': cena_za,
		           'komnati': komnat,
		           'plosh_ob':plosh_ob,
		           'etach': etash,
	                   'ettts': ets,
		           'sost': sost,
	                   'mesto': self.sub+', '+oper,
		           'rinok': rinok,
		           'opis':opis,
		           'phone':phone,
		           'lico':clearText,
		           'company':comp,
		           'data':data}
	  
	       yield Task('write',project=projects,grab=grab)
	    
	  def task_write(self,grab,task):
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['object']
	       print  task.project['cena']
	       print  task.project['komnati']
	       print  task.project['plosh_ob']
	       print  task.project['cena_za']
	       print  task.project['rayon']
	       print  task.project['etach']
	       print  task.project['ettts']
	       print  task.project['mesto']
	       print  task.project['sost']
	       print  task.project['rinok']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['dom']
	       
	       
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 17, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 40, task.project['dom'])
	       self.ws.write(self.result, 10, task.project['object'])
	       self.ws.write(self.result, 11, self.oper)
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['komnati'])
	       self.ws.write(self.result, 15, task.project['plosh_ob'])
	       self.ws.write(self.result, 16, task.project['cena_za'])
	       self.ws.write(self.result, 19, task.project['etach'])
	       self.ws.write(self.result, 20, task.project['ettts'])
	       self.ws.write(self.result, 32, task.project['mesto'])
	       self.ws.write(self.result, 28, task.project['sost'])
	       self.ws.write(self.result, 31, task.project['rinok'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, u'Росриэлт Недвижимость')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.project['phone'])
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 41, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)	       
	       print 'Ready - '+str(self.result)
	       print  self.oper
	       print '***',i+1,'/',len(l),'***'
	       print('*'*50)	       
	       self.result+= 1
	       
	       
	       #if self.result > 200:
		    #self.stop()	       
     
     
     bot = Rosreal_Kv(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     try:
	  bot.run()
     except KeyboardInterrupt:
	  pass
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
	  







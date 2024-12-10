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

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

i = 0
l= open('links/zem_prod.txt').read().splitlines()
page = l[i]
oper = u'Продажа'


while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	  
     class Rosreal_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       
	       for p in range(1,21):
		    try:
			 time.sleep(2)
			 g = Grab(timeout=50, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http') 
			 g.go(self.f)
			 self.sub = g.doc.select(u'//a[@class="a_cityp1"]').text()
			 print self.sub
			 del g
			 break
	  
		    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.sub = ''
		    
		    
	       self.workbook = xlsxwriter.Workbook(u'zem/Rosrealt_%s' % bot.sub + u'_Земля_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Rosrealt_Земля')
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
	       self.ws.write(0, 30, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
	       self.ws.write(0, 32, u"ШИРОТА_ИСХ")
	       self.ws.write(0, 33, u"ДОЛГОТА_ИСХ")	       
	       self.result= 1
     
	  def task_generator(self):
	       yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
	   
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//a[@class="nounder"][contains(@href, "uchastok")]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    if u'rosrealt' in ur :
			 yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100) 
	
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//div[@class="nolink"][contains(text(),"Следующая страница")]/ancestor::a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except DataNotFound:
		    print('*'*100)
		    print '!!!','NO PAGE NEXT','!!!'
		    print('*'*100)
		    logger.debug('%s taskq size' % self.task_queue.size())	  
	    
	    
	  
	
	
	  def task_item(self, grab, task):
	      
	       try:
		    try:
		         ray = grab.doc.select(u'//p[@class="pbig_gray"]/a[contains(text(),"район")]').text()
		    except IndexError:
			 ray = grab.doc.select(u'//p[@class="pbig_gray"]/b/a[contains(text(),"район")]').text()
	       except IndexError:
		    ray = ''          
	       try:
                    punkt= grab.doc.select(u'//div[@id="path_block"]/div[2]/a/span').text()
	       except IndexError:
		    punkt = ''
		    
	       try:
		    ter = grab.doc.select(u'//p[@class="pbig_gray"]/b[@class="blue"]/a[@class="nounder"][contains(@href, "okrug")]').text()
	       except IndexError:
		    ter =''
		    
	       try:
		    uliza = grab.doc.select(u'//a[@class="nounder"][contains(@href, "?ul=")]').text()
	       except IndexError:
		    uliza = ''
		    
	       try:
		    price = grab.doc.select(u'//b[@class="red_price"]').text()
	       except IndexError:
		    price = ''
		    
	       try:
		    cena_za = grab.doc.select(u'//b[@class="red_price"]/preceding-sibling::text()').text().replace(u'Цена за ','').replace(u'Стоимость аренды ','').replace(u'кв.м.',u'м2').replace(':','') 
	       except IndexError:
		    cena_za = ''	       	       
		    
	       
		    
	       try:
		    plosh = grab.doc.select(u'//p[contains(text(),"площадь з/у:")]').text().split(': ')[1]
	       except IndexError:
		    plosh = ''
		    
	       
	       try: 
		    categoria = grab.doc.select(u'//a[@class="nounder"][contains(@href, "?Klass=")]').text()
	       except IndexError:
		    categoria =''
	       
		    
	       try:
		    vid = grab.doc.select(u'//p[contains(text(),"назначение з/у:")]').text().split(': ')[1]
	       except IndexError:
		    vid = '' 
		    
		    
	       try:
		    ohrana = grab.doc.select(u'//p[@class="pbig_gray"]/b[@class="blue"]').text()
	       except IndexError:
		    ohrana =''
	       try:
		    gaz = grab.doc.select(u'//p[@class="pbig_gray"]/b[@class="blue"]/a[@class="nounder"][contains(@href, "dom=")]').text()
	       except IndexError:
		    gaz =''
	       try:
		    voda = re.sub(u'^.*(?=вод)','', grab.doc.select(u'//*[contains(text(), "вод")]').text())[:3].replace(u'вод',u'есть')
	       except IndexError:
		    voda =''
	       try:
		    kanal = re.sub(u'^.*(?=санузел)','', grab.doc.select(u'//*[contains(text(), "санузел")]').text())[:7].replace(u'санузел',u'есть')
	       except IndexError:
		    kanal =''
	       try:
		    elek = re.sub('[^\d\.]', u'',grab.doc.rex_text(u'ymaps.Placemark(.*?)]').split(', ')[0])
	       except IndexError:
		    elek =''
	       try:
		    teplo = re.sub('[^\d\.]', u'',grab.doc.rex_text(u'ymaps.Placemark(.*?)]').split(', ')[1])
	       except IndexError:
		    teplo =''

	       try:
		    opis = grab.doc.select(u'//div[@class="info_self"]').text()  
	       except IndexError:
		    opis = ''
		    
	       try:
		    phone = grab.doc.select(u'//p[@class="pbig_gray_contact"]/a[contains(@href, "tel:")]').text()
	       except IndexError:
		    phone = ''
		    
	       try:
		    lico = grab.doc.select(u'//p[contains(text(),"Автор объявления")]/following-sibling::p[1]').text().split(', ')[0]
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//p[contains(text(),"Автор объявления")]/following-sibling::p[1]').text().split(', ')[1]
	       except IndexError:
		    comp = ''
		    
	       try:
		    data = re.sub('[^\d\.]', u'',grab.doc.select(u'//p[contains(text(),"добавлено:")]').text())
	       except IndexError:
		    data = '' 
			 
	       try:
		    vid_prava = re.sub('[^\d\.]', u'',grab.doc.select(u'//p[contains(text(),"обновлено:")]').text())
	       except IndexError:
		    vid_prava =''
							
		    
	       projects = {'url': task.url,
		           'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'cena': price,
		           'cena_za': cena_za.replace(u' в ',u'/').replace(u'Общая стоимость',''),
		           'plosh':plosh,
		           'categoria': categoria,
		           'vid': vid,
		           'ohrana':self.sub+', '+ohrana,
		           'gaz': gaz,
		           'voda': voda,
		           'kanaliz': kanal,
		           'electr': elek,
		           'teplo': teplo,
		           'vid_prava': vid_prava,
		           'opis':opis,
		           'phone':phone,
		           'lico':lico[2:],
		           'company':comp,
		           'data':data}
	  
	       yield Task('write',project=projects,grab=grab)
	    
	  def task_write(self,grab,task):
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['cena']+task.project['cena_za']
	       print  task.project['plosh']
	       print  task.project['categoria']
	       print  task.project['vid']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['teplo']
	       print  task.project['vid_prava']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['ohrana']
	      
	       
	       
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 9, oper)
	       self.ws.write(self.result, 10, task.project['cena']+task.project['cena_za'])
	       #self.ws.write(self.result, 11, task.project['cena_za'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 13, task.project['categoria'])
	       self.ws.write(self.result, 14, task.project['vid'])
	       self.ws.write(self.result, 5, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 32, task.project['electr'])
	       self.ws.write(self.result, 33, task.project['teplo'])
	       self.ws.write(self.result, 31, task.project['ohrana'])	       
	       self.ws.write(self.result, 30, task.project['vid_prava'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'Росриэлт Недвижимость')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 28, task.project['data'])
	       self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       
	       #print task.project['koll']
	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '*',i+1,'/',len(l),'*'
	       print  oper
	       print('*'*50)	       
	       self.result+= 1
	       
	       
	       #if self.result > 20:
		    #self.stop()	       
     
     
     bot = Rosreal_Zem(thread_number=5,network_try_limit=1000)
     #bot.setup_queue('mongo', database='RosrealtZem',host='192.168.10.200')
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=500, connect_timeout=500)
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
	  if oper == u'Продажа':
	       i = 0
	       l= open('links/zem_arenda.txt').read().splitlines()
	       dc = len(l)
	       page = l[i]  
	       oper = u'Аренда'
	  else:
	       break
     






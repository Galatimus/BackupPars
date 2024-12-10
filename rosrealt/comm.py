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
l= open('links/Com_arenda.txt').read().splitlines()
page = l[i]
oper = u'Аренда'


     
while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     
     
     
     
     class Rosreal_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
     
	       for p in range(1,51):
		    try:
			 time.sleep(0.5)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http') 
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
		    
	       self.workbook = xlsxwriter.Workbook(u'com/Rosrealt_%s' % bot.sub + u'_Коммерческая_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Rosrealt_Коммерческая')
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
	       yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
     
     
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//a[@class="nounder"][contains(@href, "kommercheskaja")]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    if u'rosrealt' in ur:
			 yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	       
	       
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//div[@class="nolink"][contains(text(),"Следующая страница")]/ancestor::a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print('*'*100)
		    print '!!!','NO PAGE NEXT','!!!'
		    print('*'*100)
		       
	
	  def task_item(self, grab, task):
	      
	       try:
		    ray = grab.doc.select(u'//p[@class="pbig_gray"]/b/a[contains(text(),"район")]').text()
		  #print ray 
	       except IndexError:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//div[@id="path_block"]/div[2]/a/span').text()
	       except IndexError:
		    punkt = ''
		    
	       try:
		    try:
		         ter = grab.doc.select(u'//p[@class="pbig_gray"]/a[contains(@title, "длительно")][1]').text()
		    except IndexError:
			 ter = grab.doc.select(u'//p[@class="pbig_gray"]/a[contains(@title, "микрорайоне")][1]').text()
	       except IndexError:
		    ter =''
		    
	       try:
		    uliza = grab.doc.select(u'//a[@class="nounder"][contains(@href, "?ul=")]').text()
	       except IndexError:
		    uliza = ''
     
	       try:
		    naz = grab.doc.select(u'//p[contains(text(),"тип недвижимости:")]/a').text()
	       except IndexError:
		    naz = ''
		    
	       try:
		    price = grab.doc.select(u'//b[@class="red_price"]').text()
	       except IndexError:
		    price = ''
	       try:
		    cena_za = grab.doc.select(u'//b[@class="red_price"]/preceding-sibling::text()').text().replace(u'Цена за ','').replace(u'Стоимость аренды ','').replace(u'кв.м.',u'м2').replace(':','') 
	       except IndexError:
		    cena_za = ''	       
	       
		    
	       try:
		    plosh = grab.doc.select(u'//p[contains(text(),"общая площадь:")]').text().split(': ')[1]
	       except IndexError:
		    plosh = ''
		    
	       
	       try: 
		    klass = grab.doc.select(u'//p[@class="pbig_gray"]/a[contains(@href, "Klass=")]').text()
	       except IndexError:
		    klass =''
	       try:
	            ohrana = grab.doc.select(u'//p[@class="pbig_gray"]/b[@class="blue"]').text()
	       except IndexError:
		    ohrana =''
	       try:
		    gaz = grab.doc.select(u'//p[@class="pbig_gray"][1]/b/a[contains(@href, "ul=")]/following-sibling::text()').text().replace(', ','').replace('/','|')
	       except IndexError:
		    gaz =''
	       try:
		    voda = grab.doc.select(u'//div[@class="section_right"]/p[2]').text()
	       except IndexError:
		    voda =''
	       try:
		    kanal = grab.doc.select(u'//font[@color="#EB1E01"]').text()
	       except IndexError:
		    kanal =''
	       try:
		    elek =  re.sub('[^\d\.]', u'',grab.doc.rex_text(u'ymaps.Placemark(.*?)]').split(', ')[0])
	       except IndexError:
		    elek =''
		    
	       try:
	            lng =  re.sub('[^\d\.]', u'',grab.doc.rex_text(u'ymaps.Placemark(.*?)]').split(', ')[1])
	       except IndexError:
	            lng =''
		    
	       try:
		    teplo = grab.doc.select(u'//h1').text()
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
		    data = re.sub('[^\d\.]', u'',grab.doc.select(u'//p[contains(text(),"обновлено:")]').text())
	       except IndexError:
	            data = '' 
		    
	       try:
	            data1 = re.sub('[^\d\.]', u'',grab.doc.select(u'//p[contains(text(),"добавлено:")]').text())
	       except IndexError:
                    data1 = ''		    
			 
	       
							
		    
	       projects = {'url': task.url,
		           'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter.replace(naz,''),
		           'ulica': uliza,
		           'naz': naz,
		           'cena': price,
		           'cena_za': cena_za.replace(u' в ',u'/').replace(u'Общая стоимость',''),
		           'plosh':plosh,
		           'klass': klass,
		           'ohrana':self.sub+', '+ohrana,
	                   'voda':voda,
		           'gaz': gaz.replace(',',''),
		            'kanaliz': kanal,
		           'electr': elek,
	                   'dol': lng,
		           'teplo': teplo,
		           'opis':opis,
		           'phone':phone,
		           'lico':lico[2:],
		           'company':comp,
	                   'datad':data1,
		           'data':data}
	  
	       yield Task('write',project=projects,grab=grab)
	    
	  def task_write(self,grab,task):
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['naz']
	       print  task.project['cena']+task.project['cena_za']
	       print  task.project['plosh']
	       print  task.project['klass']
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
	       print  task.project['ohrana']
	       print  task.project['datad']
	       
	       
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       self.ws.write(self.result, 28, oper)
	       self.ws.write(self.result, 11, task.project['cena']+task.project['cena_za'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 10, task.project['klass'])
	       self.ws.write(self.result, 5, task.project['gaz'])
	       self.ws.write(self.result, 25, task.project['voda'])
	       self.ws.write(self.result, 26, task.project['kanaliz'])
	       self.ws.write(self.result, 34, task.project['electr'])
	       self.ws.write(self.result, 35, task.project['dol'])
	       self.ws.write(self.result, 33, task.project['teplo'])
	       self.ws.write(self.result, 24, task.project['ohrana'])	       
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'Росриэлт Недвижимость')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 29, task.project['datad'])
	       self.ws.write(self.result, 30, task.project['data'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       
	       print 'Ready - '+str(self.result)
	       print  oper
	       print '***',i+1,'/',len(l),'***'
	       print('*'*50)	       
	       self.result+= 1
	       
	       
		  
     
     
     bot = Rosreal_Com(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
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
          if oper == u'Аренда':
               i = 0
               l= open('links/Com_prod.txt').read().splitlines()
               page = l[i]
               oper = u'Продажа'
          else:
	       break








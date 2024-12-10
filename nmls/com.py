#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError,GrabTooManyRedirectsError 
import logging
import re
import time
import os
from grab import Grab
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
    class MK_Com(Spider):
	def prepare(self):
	    self.f = page
	    self.link =l[i]
	    for p in range(1,31):
	        try:
                    time.sleep(1)
		    g = Grab(timeout=20, connect_timeout=50)
		    g.proxylist.load_file(path='../ivan.txt',proxy_type='http') 
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
		    
	    self.workbook = xlsxwriter.Workbook(u'com/Nmls_%s' % bot.sub + u'_Коммерческая_'+str(i)+'.xlsx')
	    self.ws = self.workbook.add_worksheet()
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
		print('*'*100)
		print '!!!!!','NO PAGE NEXT','!!!!'
		print('*'*100)

	def task_item(self, grab, task):
	    
	    
	    try:
		ray = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/span/a[contains(@href,"raions")]').text().replace(u'район','')
	    except DataNotFound:
		ray =''
	    try:
		punkt= grab.doc.select(u'//td[contains(text(),"Город")]/following-sibling::td').text().replace(u'Область','')
	    except IndexError:
		punkt = ''
	    try:
		try:
		    ter = grab.doc.select(u'//td[contains(text(),"Район")]/following-sibling::td').text()
		except DataNotFound:
		    ter = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/span/a[contains(text(),"поселок")]').text()
		#else:
		    #ter= ''
	    except IndexError:
		ter =''    
	    try:
		try:
		    try:
		        uliza = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/a[contains(text(),"улица")]').text()
		    except DataNotFound:
		        uliza = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/a[contains(text(),"переулок")]').text()
		except DataNotFound:
		    uliza = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td/a[contains(text(),"проспект")]').text()
	    except DataNotFound:
		uliza = '' 
	    try:
		dom = grab.doc.select(u'//label[contains(text(),"Адрес:")]/following-sibling::p/a[contains(@href,"houseId")]').text()
	    except DataNotFound:
		dom = ''
		
	    try:
		orentir = grab.doc.select(u'//td[contains(text(),"Материал")]/following-sibling::td').text()
	    except DataNotFound:
		orentir = ''
		
	    try:
	        seg = grab.doc.select(u'//td[contains(text(),"Объект")]/following-sibling::td').text()
	      #print oren
	    except DataNotFound:
		seg = '' 
		
	    try:
	        naz = grab.doc.select(u'//td[contains(text(),"Использование")]/following-sibling::td').text()
	      #print naz
	    except DataNotFound:
		naz = '' 
		
	    try:
		klass = grab.doc.select(u'//h1').text()
	    except DataNotFound:
		klass = ''
		
	    try:
	        price = grab.doc.select(u'//div[@class="card-price"]').text()
	      #print price
	    except DataNotFound:
		price = ''
		
	    try:
	        plosh = grab.doc.select(u'//td[contains(text(),"Площадь здания/земли")]/following-sibling::td').text().split('/')[0]+u' м2'
	      #print plosh
	    except DataNotFound:
		plosh = '' 
		
	    try:
		et = grab.doc.select(u'//label[contains(text(),"Этаж:")]/following-sibling::p').text().split(u'из ')[0]
	    except IndexError:
		et = ''
		
	    try:
		et2 = grab.doc.select(u'//td[contains(text(),"Дата подачи объявления")]/following-sibling::td').text().split(u' в ')[0]
	    except IndexError:
		et2 = ''
		
	    try:
	        opis = grab.doc.select(u'//div[@class="object-infoblock"][2]').text()
	      #print opis
	    except DataNotFound:
		opis = ''
		
	    try:
		phone = re.sub('[^\d\,]','',grab.doc.select(u'//div[@class="mb5"]/img[contains(@src,"phone")]/following-sibling::text()[1]').text())
	      #print phone
	    except DataNotFound:
		phone = '' 
		
	    try:
		lico =  re.sub('[\d]','',grab.doc.select(u'//div[contains(text(),"Агент:")]/a').text())#.split(', ')[1]
	    except IndexError:
		lico = ''
		
	    try:
	        comp = grab.doc.select(u'//div[contains(text(),"АН:")]/a').text()
	      #print comp
	      
	    except DataNotFound:
		comp = '' 
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
		teplo = grab.doc.select(u'//div[@id="YMapsID"]').attr('data-geocode')
	    except DataNotFound:
		teplo =''
		
	    try:
		data= grab.doc.select(u'//td[contains(text(),"Дата обновления")]/following-sibling::td').text().split(u' в ')[0]
	    except IndexError:
		data = ''
		
	    try:
		if 'prodazha' in task.url:
	            oper = u'Продажа' 
	        elif 'arenda' in task.url:
		    oper = u'Аренда'     
            except IndexError:
	        oper = ''	    
	    
	    projects = {'url': task.url,
		        'sub': self.sub,
		        'ray': ray,
		        'punkt': punkt,
		        'teritor': ter,
		        'uliza': uliza,
		        'dom': dom,
		        'orentir':orentir,
		        'seg': seg,
		        'naznachenie': naz,
		        'klass': klass,
		        'cena': price,
		        'ploshad': plosh,
	                'oper':oper,
		        'et': et,
		        'ets': et2,
		        'opisanie': opis,
		        'phone':phone,
		        'company':comp,
		        'lico':lico,
		        'ohrana':ohrana,
		        'gaz': gaz,
		        'voda': voda,
		        'kanaliz': kanal,
		        'electr': elek,
		        'teplo': teplo,
		        'data':data}
	    
	    yield Task('write',project=projects,grab=grab)
	    
	def task_write(self,grab,task):
	    if task.project['opisanie'] <> '':
		print('*'*50)
		print  task.project['sub']
		print  task.project['ray']
		print  task.project['punkt']
		print  task.project['teritor']
		print  task.project['uliza']
		print  task.project['dom']
		print  task.project['orentir']
		print  task.project['seg']
		print  task.project['naznachenie']
		print  task.project['klass']
		print  task.project['cena']
		print  task.project['ploshad']
		print  task.project['et']
		print  task.project['opisanie']
		print  task.project['url']
		print  task.project['phone']
		print  task.project['lico']
		print  task.project['company']
		print  task.project['ohrana']
		print  task.project['gaz']
		print  task.project['voda']
		print  task.project['kanaliz']
		print  task.project['electr']
		print  task.project['data']
		print  task.project['ets']
		print  task.project['teplo']
		
		
		
		self.ws.write(self.result, 0, task.project['sub'])
		self.ws.write(self.result, 1, task.project['ray'])
		self.ws.write(self.result, 2, task.project['punkt'])
		self.ws.write(self.result, 3, task.project['teritor'])
		self.ws.write(self.result, 4, task.project['uliza'])
		self.ws.write(self.result, 5, task.project['dom'])
		#self.ws.write(self.result, 16, task.project['orentir'])
		self.ws.write(self.result, 7, task.project['seg'])
		#self.ws.write(self.result, 8, task.project['tip'])
		self.ws.write(self.result, 9, task.project['naznachenie'])
		self.ws.write(self.result, 33, task.project['klass'])
		self.ws.write(self.result, 11, task.project['cena'])
		self.ws.write(self.result, 14, task.project['ploshad'])	
		#self.ws.write(self.result, 13, task.project['et'])
		self.ws.write(self.result, 29, task.project['ets'])
		#self.ws.write(self.result, 15, task.project['god'])
		#self.ws.write(self.result, 16, task.project['mat'])
		#self.ws.write(self.result, 17, task.project['potolok'])
		#self.ws.write(self.result, 18, task.project['sost'])
		#self.ws.write(self.result, 19, task.project['ohrana'])
		#self.ws.write(self.result, 20, task.project['gaz'])
		#self.ws.write(self.result, 21, task.project['voda'])
		#self.ws.write(self.result, 22, task.project['kanaliz'])
		self.ws.write(self.result, 28, task.project['oper'])
		self.ws.write(self.result, 24, task.project['teplo'])
		self.ws.write_string(self.result, 20, task.project['url'])
		self.ws.write(self.result, 21, task.project['phone'])
		self.ws.write(self.result, 22, task.project['lico'])
		self.ws.write(self.result, 23, task.project['company'])
		self.ws.write(self.result, 30, task.project['data'])
		self.ws.write(self.result, 18, task.project['opisanie'])
		self.ws.write(self.result, 19, u'Система "ИС Центр"')
		self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
		#self.ws.write(self.result, 28, oper)
		
		
		print('*'*50)
		print self.sub
		print 'Ready - '+str(self.result)
		print '***',i+1,'/',dc,'***'
		print task.project['oper']
		print('*'*50)
		
		self.result+= 1
		
		#if self.result > 10:
		    #self.stop()	

    bot = MK_Com(thread_number=5,network_try_limit=1000)
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
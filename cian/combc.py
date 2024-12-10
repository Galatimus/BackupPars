#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import math
import re
import os
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)



workbook = xlsxwriter.Workbook(u'CIAN_Объекты.xlsx')



class Cian_Com(Spider):
    def prepare(self):
	self.ws = workbook.add_worksheet()
	self.ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	self.ws.write(0, 1, u"НАСЕЛЕННЫЙ_ПУНКТ")
	self.ws.write(0, 2, u"АДРЕС")
	self.ws.write(0, 3, u"СЕГМЕНТ")
	self.ws.write(0, 4, u"НАИМЕНОВАНИЕ")
	self.ws.write(0, 5, u"КЛАСС")
	self.ws.write(0, 6, u"ОБЩАЯ_ПЛОЩАДЬ")
	self.ws.write(0, 7, u"ГОД_ПОСТРОЙКИ")
	self.ws.write(0, 8, u"ЭТАЖНОСТЬ")
	self.ws.write(0, 9, u"ПАРКОВКА")
	self.ws.write(0, 10, u"НАПОЛЬНОЕ_ПОКРЫТИЕ")
	self.ws.write(0, 11, u"КРАНОВОЕ_ОБОРУДОВАВНИЕ")
	self.ws.write(0, 12, u"ВОРОТА")
	self.ws.write(0, 13, u"ТЕМПЕРАТУРНЫЙ РЕЖИМ")
	self.ws.write(0, 14, u"ПЛОЩАДЬ ОТКРЫТОЙ ПЛОЩАДКИ")
	self.ws.write(0, 15, u"ПОТОЛКИ")
	self.ws.write(0, 16, u"ШАГ КОЛОНН")
	self.ws.write(0, 17, u"СТЕНЫ И НЕСУЩИЕ КОСТРУКЦИИ")
	self.ws.write(0, 18, u"КОНДИЦИОНИРОВАНИЕ")
	self.ws.write(0, 19, u"ЛИФТ")
	self.ws.write(0, 20, u"ДОСТУП В ЗДАНИЕ")
	self.ws.write(0, 21, u"ВХОД")
	self.ws.write(0, 22, u"ИНФРАСТРУКТУРА")
	self.ws.write(0, 23, u"АРЕНДАТОРЫ")
	self.ws.write(0, 24, u"МЕТЕРИАЛ ПЕРЕКРЫТИЙ")
	self.ws.write(0, 25, u"ВЕНТИЛЯЦИЯ")
	self.ws.write(0, 26, u"ПОЖАРОТУШЕНИЕ")
	self.ws.write(0, 27, u"СИСТЕМА Б/П")
	self.ws.write(0, 28, u"КАТЕГОРИЯ ЕЛЕКТРОНАДЕЖНОСТИ")
	self.ws.write(0, 29, u"АРЕНДНАЯ СТАВКА")
	self.ws.write(0, 30, u"СТОИМОСТЬ")
	self.ws.write(0, 31, u"АРЕНДОПРИГОДНАЯ ПЛОЩАДЬ")
	self.ws.write(0, 32, u"ОПИСАНИЕ")
	self.ws.write(0, 33, u"ПРОДАВЕЦ")
	self.ws.write(0, 34, u"ТЕЛЕФОН")
	self.ws.write(0, 35, u"ССЫЛКА")
	self.ws.write(0, 36, u"ДАТА ПАРСИНГА")	
	self.result= 1
	#self.count = 2
	
	    
	    
	    
	      
    
    def task_generator(self):
	l= open('cian_bc.txt').read().splitlines()
	self.dc = len(l)
	print self.dc
	for line in l:
	    yield Task ('item',url=line,network_try_count=100)

    def task_item(self, grab, task):
	#time.sleep(1)
	
	try:
	    sub = grab.doc.select(u'//address[contains(@class,"address")]').text().split(', ')[0]
	except IndexError:
	    sub = ''	
	try:
	    usl = grab.doc.select(u'//h2').text().replace(u'Про ','')
	except IndexError:
	    usl = ''	
	try:
	    ray = grab.doc.select(u'//address[contains(@class,"address")]/p[1]').text().replace(u'На карте','')
	except IndexError:
	    ray =''
	try:
	    if sub == u'Москва':
		punkt= u'Москва'
	    elif sub == u'Санкт-Петербург':
		punkt= u'Санкт-Петербург'
	    elif sub == u'Севастополь':
		punkt= u'Севастополь'
	    else:
		if  grab.doc.select(u'//address[contains(@class,"address")]/a[2][contains(text(),"р-н ")]').exists()==True:
		    punkt= grab.doc.select(u'//address[contains(@class,"address")]/a[3]').text()
		elif grab.doc.select(u'//address[contains(@class,"address")]/a[3][contains(text(),"р-н ")]').exists()==True:
		    punkt= grab.doc.select(u'//address[contains(@class,"address")]/a[2]').text()
		else:
		    punkt=grab.doc.select(u'//address[contains(@class,"address")]/a[2]').text()
	except IndexError:
	    punkt = ''
	try:
	    ter=  grab.doc.select(u'//h1').text()
	except IndexError:
	    ter =''
	try:
	    uliza = grab.doc.select(u'//dt[contains(text(),"Класс")]/following-sibling::dd[1]').text()
	except IndexError:
	    uliza =''
	
	try:
	    dom = grab.doc.select(u'//dt[contains(text(),"Общая площадь")]/following-sibling::dd[1]').text()
	except IndexError:
	    dom = ''
	    
	try:
	    seg = grab.doc.select(u'//dt[contains(text(),"Парковка")]/following-sibling::dd[1]/p').text()
	  #print oren
	except IndexError:
	    seg = '' 
	    
	try:
	    naz = grab.doc.select(u'//dt[contains(text(),"Год постройки")]/following-sibling::dd[1]').text()
	  #print naz
	except IndexError:
	    naz = '' 
	    
	try:
	    klass = grab.doc.select(u'//dt[contains(text(),"Этажность")]/following-sibling::dd[1]').text()
	except IndexError:
	    klass = ''
	    
	try:
	    price = grab.doc.select(u'//dt[contains(text(),"Покрытие пола")]/following-sibling::dd[1]').text()
	  #print price
	except IndexError:
	    price = ''
	    
	try:
	    plosh = grab.doc.select(u'//dt[contains(text(),"Крановое оборудование")]/following-sibling::dd[1]').text()
	  #print plosh
	except IndexError:
	    plosh = '' 
	    
	try:
	    et = grab.doc.select(u'//dt[contains(text(),"Ворота")]/following-sibling::dd[1]').text()
	except IndexError:
	    et = ''
	    
	try:
	    et2 = grab.doc.select(u'//dt[contains(text(),"Температурный режим")]/following-sibling::dd[1]').text()
	except IndexError:
	    et2 = ''
	    
	try:
	    opis = grab.doc.select(u'//dt[contains(text(),"Площадь открытой площадки")]/following-sibling::dd[1]').text()
	  #print opis
	except IndexError:
	    opis = ''
	    
	try:
	    phone = grab.doc.select(u'//dt[contains(text(),"Потолки")]/following-sibling::dd[1]').text()
	except IndexError:
	    phone = '' 
	    
	try:
	    lico = grab.doc.select(u'//dt[contains(text(),"Сетка колонн")]/following-sibling::dd[1]').text()
	except IndexError:
	    lico = ''
	    
	try:
	    comp = grab.doc.select(u'//dt[contains(text(),"Стены и несущие конструкции")]/following-sibling::dd[1]').text()
	except IndexError:
	    comp = '' 
	    
	try:
	    ohrana = grab.doc.select(u'//dt[contains(text(),"Кондиционирование")]/following-sibling::dd[1]').text()
	except IndexError:
	    ohrana =''
	try:
	    gaz = grab.doc.select(u'//dt[contains(text(),"Лифт")]/following-sibling::dd[1]').text()
	except IndexError:
	    gaz =''
	try:
	    voda =  grab.doc.select(u'//dt[contains(text(),"Доступ в здание")]/following-sibling::dd[1]').text()
	except IndexError:
	    voda =''
	try:
	    kanal = grab.doc.select(u'//dt[contains(text(),"Вход")]/following-sibling::dd[1]').text()
	except IndexError:
	    kanal =''
	try:
	    elek = grab.doc.select(u'//dt[contains(text(),"Инфраструктура")]/following-sibling::dd[1]').text()
	except IndexError:
	    elek =''
	try:
	    teplo = grab.doc.select(u'//h2[contains(text(),"Арендаторы")]/following-sibling::div[1]/p').text()
	except IndexError:
	    teplo =''
	    
	try:
            data = grab.doc.select(u'//dt[contains(text(),"Материал перекрытий")]/following-sibling::dd[1]').text()
	except IndexError:
	    data = ''
	    
	    
	try:
	    lat = grab.doc.select(u'//dt[contains(text(),"Вентиляция")]/following-sibling::dd[1]').text()
	except IndexError:
	    lat =''
    
	try:
	    lng = grab.doc.select(u'//dt[contains(text(),"Пожаротушение")]/following-sibling::dd[1]').text()
	except IndexError:
	    lng =''
	    
	try:
	    cond = grab.doc.select(u'//dt[contains(text(),"Система бесперебойного питания")]/following-sibling::dd[1]').text()
	except IndexError:
	    cond =''
	    
	try:
	    vent = grab.doc.select(u'//dt[contains(text(),"Категория электронадежности")]/following-sibling::dd[1]').text()
	except IndexError:
	    vent =''		
	    
	try:
	    uslu = grab.doc.select(u'//dt[contains(text(),"Арендная ставка")]/following-sibling::dd[1]').text()
	except IndexError:
	    uslu = ''
	    
	try:
            oper = grab.doc.select(u'//dt[contains(text(),"Стоимость")]/following-sibling::dd[1]').text()     
	except IndexError:
	    oper = ''
	try:
	    oper1 = grab.doc.select(u'//dt[contains(text(),"Арендопригодная площадь")]/following-sibling::dd[1]').text()     
	except IndexError:
	    oper1 = ''	    
	try:
	    oper2 = grab.doc.select(u'//div[contains(@class,"description_limit")]/p').text()     
	except IndexError:
	    oper2 = ''	
        try:
	    oper3 = grab.doc.rex_text(u'href="tel:(.*?)"')     
	except IndexError:
	    oper3 = ''
	    
	    
	projects = {'url': task.url,
                    'sub': sub,
                    'ray': ray,
                    'punkt': punkt.replace(u' городской округ',''),
                    'teritor': ter,
                    'uliza': uliza,
                    'dom': dom,
                    'seg': seg,
                    'naznachenie': naz,
                    'klass': klass,
                    'uslovi': usl,
                    'uslugi':uslu,
                    'cena': price,
                    'ploshad': plosh,
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
	            'oper1':oper1,
	            'oper2':oper2,
	            'oper3':oper3,
                    'teplo': teplo,
                    'condi':cond,
                    'internet':vent,
                    'dol': lat,
                    'shir': lng,	                
                    'data':data,
                    'oper':oper
                    
                    }
	yield Task('write',project=projects,grab=grab)
	
    def task_write(self,grab,task):
	if task.project['sub'] <> '':    
	    print('*'*50)
	    print  task.project['sub']
	    print  task.project['ray']
	    print  task.project['punkt']
	    print  task.project['teritor']
	    print  task.project['uliza']
	    print  task.project['dom']
	    print  task.project['seg']
	    print  task.project['naznachenie']
	    print  task.project['oper2']
	    print  task.project['uslovi'] 
	    print  task.project['klass']
	    print  task.project['cena']
	    print  task.project['ploshad']
	    print  task.project['et']
	    print  task.project['ets']
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
	    print  task.project['teplo']
	    print  task.project['data']
	    
	    
	    
	    
	    self.ws.write(self.result, 0, task.project['sub'])
	    self.ws.write(self.result, 2, task.project['ray'])
	    self.ws.write(self.result, 1, task.project['punkt'])
	    self.ws.write(self.result, 3, task.project['uslovi'])
	    self.ws.write(self.result, 4, task.project['teritor'])
	    self.ws.write(self.result, 5, task.project['uliza'])
	    self.ws.write(self.result, 6, task.project['dom'])
	    self.ws.write(self.result, 7, task.project['naznachenie'])
	    self.ws.write(self.result, 28, task.project['internet'])
	    self.ws.write(self.result, 8, task.project['klass'])
	    self.ws.write(self.result, 9, task.project['seg'])
	    self.ws.write(self.result, 29, task.project['uslugi']) 
	    self.ws.write(self.result, 10, task.project['cena'])
	    self.ws.write(self.result, 11, task.project['ploshad'])  
	    self.ws.write(self.result, 12, task.project['et'])
	    self.ws.write(self.result, 13, task.project['ets'])
	    self.ws.write(self.result, 14, task.project['opisanie'])
	    self.ws.write(self.result, 15, task.project['phone'])
	    self.ws.write(self.result, 27, task.project['condi'])
	    self.ws.write(self.result, 26, task.project['shir'])
	    self.ws.write(self.result, 25, task.project['dol'])
	    self.ws.write(self.result, 18, task.project['ohrana'])
	    self.ws.write(self.result, 19, task.project['gaz'])
	    self.ws.write(self.result, 20, task.project['voda'])
	    self.ws.write(self.result, 21, task.project['kanaliz'])
	    self.ws.write(self.result, 22, task.project['electr'])
	    self.ws.write(self.result, 23, task.project['teplo'])
	    self.ws.write_string(self.result, 35, task.project['url'])
	    self.ws.write(self.result, 16, task.project['lico'])
	    self.ws.write(self.result, 17, task.project['company'])
	    self.ws.write(self.result, 24, task.project['data'])  
	    #self.ws.write(self.result, 19, u'ЦИАН')
	    self.ws.write(self.result, 36, datetime.today().strftime('%d.%m.%Y'))
	    self.ws.write(self.result, 30, task.project['oper'])
	    self.ws.write(self.result, 31, task.project['oper1'])
	    self.ws.write(self.result, 32, task.project['oper2'])
	    self.ws.write(self.result, 34, task.project['oper3'])
	    
	    
	    print('*'*50)
	    print 'Ready - '+str(self.result)+'/'+str(self.dc)
	    print 'Tasks - %s' % self.task_queue.size()
	    print  task.project['oper']
	    print('*'*50)
	    
	    self.result+= 1
	    
	    
	    
	    #if self.result > 10:
		#self.stop()	
		
	    
	   
bot = Cian_Com(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
try:
    bot.run()
except KeyboardInterrupt:
    pass
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
command = 'mount -a'
os.system('echo %s|sudo -S %s' % ('1122', command))
time.sleep(3)
workbook.close()
print('Done')



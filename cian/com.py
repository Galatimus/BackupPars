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



workbook = xlsxwriter.Workbook(u'0001-0002_00_C_005-0002_CIAN.xlsx')



class Cian_Com(Spider):
    def prepare(self):
	self.ws = workbook.add_worksheet(u'Cian')
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
	self.ws.write(0, 36, u"ТРАССА")
	self.ws.write(0, 37, u"ПАРКОВКА")
	self.ws.write(0, 38, u"ОХРАНА")
	self.ws.write(0, 39, u"КОНДИЦИОНИРОВАНИЕ")
	self.ws.write(0, 40, u"ИНТЕРНЕТ")
	self.ws.write(0, 41, u"ТЕЛЕФОН (КОЛИЧЕСТВО ЛИНИЙ)")
	self.ws.write(0, 42, u"УСЛУГИ")
	self.ws.write(0, 43, u"СИСТЕМА ВЕНТИЛЯЦИИ")    
	self.result= 1
	#self.count = 2
	
	    
	    
	    
	      
    
    def task_generator(self):
	l= open('cian_com.txt').read().splitlines()
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
	    usl = grab.doc.select(u'//div[contains(@class,"price_changes")]/following-sibling::p').text()
	except IndexError:
	    usl = ''	
	try:
	    ray = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"р-н ")]').text()
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
	    ter=  grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"городской округ")]').text()
	except IndexError:
	    ter =''
	try:
	    try:
		try:
		    try:
			try:
			    try:
				try:
				    try:
					uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"ул.")]').text()
				    except IndexError:
					uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"пер.")]').text()
				except IndexError:
				    uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"просп.")]').text()
			    except IndexError:
				uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"ш.")]').text()
			except IndexError:
			    uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"бул.")]').text()
		    except IndexError:
			uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"проезд")]').text()
		except IndexError:
		    uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"наб.")]').text()
	    except IndexError:
		uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"пл.")]').text()
	except IndexError:
	    uliza =''
	
	try:
	    if uliza == '':
		dom =''
	    else:
		dom = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(@href,"house")]').text()
	except IndexError:
	    dom = ''
	    
	try:
	    seg = grab.doc.select(u'//span[contains(text(),"Тип здания")]/following-sibling::span[1]').text()
	  #print oren
	except DataNotFound:
	    seg = '' 
	    
	try:
	    naz = grab.doc.select(u'//h1').text().split(', ')[0].split('(')[0]
	  #print naz
	except IndexError:
	    naz = '' 
	    
	try:
	    klass = grab.doc.select(u'//div[contains(text(),"Класс")]/following-sibling::div[1]').text()
	except IndexError:
	    klass = ''
	    
	try:
	    price = grab.doc.select(u'//span[@itemprop="price"]').text()
	  #print price
	except IndexError:
	    price = ''
	    
	try:
	    plosh = grab.doc.select(u'//div[contains(text(),"Площадь")]/following-sibling::div[1]').text()#.replace(u'м',u'м2')
	  #print plosh
	except IndexError:
	    plosh = '' 
	    
	try:
	    et = grab.doc.select(u'//div[contains(text(),"Этаж")]/following-sibling::div[contains(text()," из ")]').text().split(u' из ')[0]
	except IndexError:
	    et = ''
	    
	try:
	    try:
	        et2 = grab.doc.select(u'//div[contains(text(),"Этаж")]/following-sibling::div[contains(text()," из ")]').text().split(u' из ')[1]
	    except IndexError:
		et2 = grab.doc.select(u'//div[contains(text(),"Этажность")]/following-sibling::div[contains(text()," этаж")]').text()
	except IndexError:
	    et2 = ''
	    
	try:
	    opis = grab.doc.select(u'//p[@itemprop="description"]').text()
	  #print opis
	except IndexError:
	    opis = ''
	    
	try:
		try:
		    phone = re.sub(u'[^\d\+]','',grab.doc.select(u'//div[@class="cf_offer_show_phone-number"]/a').text())
		except IndexError:
		    phone = re.sub(u'[^\d\+]','',grab.doc.rex_text(u'offerPhone(.*?),'))
	except IndexError:
	    phone = '' 
	    
	try:
	    try:
		lico = grab.doc.select(u'//a[contains(@href,"agents")]/h2').text()
	    except IndexError:
		lico = grab.doc.select(u'//h3[@class="realtor-card__title"]/a[contains(@href,"agents")]').text() 
	except IndexError:
	    lico = ''
	    
	try:
	    try:
		comp = grab.doc.select(u'//a[contains(@href,"company")]/h2').text()
	    except IndexError:
		comp = grab.doc.select(u'//h4[@class="realtor-card__subtitle"]').text()
	except IndexError:
	    comp = '' 
	try:
	    ohrana = grab.doc.select(u'//span[contains(text(),"Год постройки")]/following-sibling::span[1]').text()
	except IndexError:
	    ohrana =''
	try:
	    gaz = grab.doc.select(u'//address[contains(@class,"address")]').text().replace(u'На карте','')
	except IndexError:
	    gaz =''
	try:
	    voda =  grab.doc.select(u'//a[contains(@href,"metro")]/span').text()
	except IndexError:
	    voda =''
	try:
	    kanal = grab.doc.select(u'//span[contains(@class,"underground")]').text().replace(', ','')
	except IndexError:
	    kanal =''
	try:
	    elek = grab.doc.select(u'//title').text()
	except IndexError:
	    elek =''
	try:
	    teplo = grab.doc.select(u'//span[contains(text(),"Парковка")]/following-sibling::span[1]').text()
	except IndexError:
	    teplo =''
	    
	try:
            data = re.sub(u'[^\d\-]','',grab.doc.rex_text(u'editDate(.*?)T')).replace('-','.')
		   #print data
	except IndexError:
	    data = ''
	    
	    
	try:
	    lat = grab.doc.rex_text(u'center=(.*?)&').split('%2C')[0]
	except IndexError:
	    lat =''
    
	try:
	    lng = grab.doc.rex_text(u'center=(.*?)&').split('%2C')[1]
	except IndexError:
	    lng =''
	    
	try:
	    cond = grab.doc.select(u'//div[contains(text(),"Центральное кондиционирование")]').text()
	except IndexError:
	    cond =''
	    
	try:
	    vent = grab.doc.select(u'//div[contains(text(),"Приточная вентиляция")]').text()
	except IndexError:
	    vent =''		
	    
	try:
	    li = []
	    for e in grab.doc.select(u'//ul[@class="cf-comm-offer-detail__infrastructure"]/li'):
		ur = e.text()
		#print ur
		li.append(ur)		
	    uslu = ",".join(li)
	except IndexError:
	    uslu = ''
	    
	try:
	    if 'sale' in task.url:
	        oper = u'Продажа' 
	    elif 'rent' in task.url:
	        oper = u'Аренда'     
	except IndexError:
	    oper = ''
	
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
                    'phone':phone.replace(u'79311111111',''),
                    'company':comp,
                    'lico':lico,
                    'ohrana':ohrana,
                    'gaz': gaz,
                    'voda': voda,
                    'kanaliz': kanal,
                    'electr': elek,
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
	    self.ws.write(self.result, 1, task.project['ray'])
	    self.ws.write(self.result, 2, task.project['punkt'])
	    self.ws.write(self.result, 3, task.project['teritor'])
	    self.ws.write(self.result, 4, task.project['uliza'])
	    self.ws.write(self.result, 5, task.project['dom'])
	    self.ws.write(self.result, 43, task.project['internet'])
	    self.ws.write(self.result, 8, task.project['seg'])
	    self.ws.write(self.result, 42, task.project['uslugi'])
	    self.ws.write(self.result, 9, task.project['naznachenie'])
	    self.ws.write(self.result, 10, task.project['klass'])
	    self.ws.write(self.result, 11, task.project['cena'])
	    self.ws.write(self.result, 14, task.project['ploshad'])	
	    self.ws.write(self.result, 13, task.project['uslovi'])
	    self.ws.write(self.result, 15, task.project['et'])
	    self.ws.write(self.result, 16, task.project['ets'])
	    self.ws.write(self.result, 39, task.project['condi'])
	    self.ws.write(self.result, 35, task.project['shir'])
	    self.ws.write(self.result, 34, task.project['dol'])
	    self.ws.write(self.result, 17, task.project['ohrana'])
	    self.ws.write(self.result, 24, task.project['gaz'])
	    self.ws.write(self.result, 26, task.project['voda'])
	    self.ws.write(self.result, 27, task.project['kanaliz'])
	    self.ws.write(self.result, 33, task.project['electr'])
	    self.ws.write(self.result, 37, task.project['teplo'])
	    self.ws.write_string(self.result, 20, task.project['url'])
	    self.ws.write(self.result, 21, task.project['phone'])
	    self.ws.write(self.result, 22, task.project['lico'])
	    self.ws.write(self.result, 23, task.project['company'])
	    self.ws.write(self.result, 30, task.project['data'])
	    self.ws.write(self.result, 18, task.project['opisanie'])
	    self.ws.write(self.result, 19, u'ЦИАН')
	    self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	    self.ws.write(self.result, 28, task.project['oper'])
	    
	    
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



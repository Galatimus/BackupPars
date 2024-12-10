#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
import xlsxwriter
import time
import os
import random
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'comm/0001-0014_00_C_001-0023_DOSKA.xlsx')




class Cian_Com(Spider):
    
    
    
    def prepare(self):
	 
	
	self.ws = workbook.add_worksheet('doska')
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
	self.result= 1

    
    def task_generator(self):
        for line in open('links/com.txt').read().splitlines():
            yield Task ('post',url=line.strip(),refresh_cache=True,network_try_count=100)
	#yield Task ('post',url='http://www.cian.ru/snyat-pomeshenie/',refresh_cache=True,network_try_count=100)
	
	
    def task_page(self,grab,task):
        try:
            pg = grab.doc.select(u'//button[@class="navia"]/following-sibling::a[contains(@href,"page")][1]')
            u = grab.make_url_absolute(pg.attr('href'))
            yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
        except DataNotFound:
            print('*'*100)
            print '!!!','NO PAGE NEXT','!!!'
            print('*'*100)
            logger.debug('%s taskq size' % self.task_queue.size())	
        
    def task_post(self,grab,task):
        for elem in grab.doc.select(u'//div[@class="d1"]/a'):
	    ur = grab.make_url_absolute(elem.attr('href'))  
	    #print ur
	    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	yield Task("page", grab=grab,refresh_cache=True,network_try_count=100,use_proxylist=False)
	            
  
	
        
        
        
    def task_item(self, grab, task):
	
        try:
            sub = grab.doc.select(u'//td[contains(text(),"Область:")]/following-sibling::td').text()#.split(', ')[0]
        except IndexError:
            sub = ''	
	try:
	    try:
                ray = grab.doc.select(u'//td[contains(text(),"айон")]/following-sibling::td/b[contains(text(),"район")]').text()
	    except IndexError:
	        ray = grab.doc.select(u'//td[contains(text(),"айон")]/following-sibling::td[contains(text(),"район")]').text()
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
		if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[2][contains(text(),"район")]').exists()==True:
                    punkt= grab.doc.select(u'//h1[@class="object_descr_addr"]/a[3]').text()
		elif grab.doc.select(u'//h1[@class="object_descr_addr"]/a[3][contains(text(),"район")]').exists()==True:
                    punkt= grab.doc.select(u'//h1[@class="object_descr_addr"]/a[2]').text()
		else:
		    punkt=grab.doc.select(u'//td[contains(text(),"Город")]/following-sibling::td').text().replace(ray,'')
        except IndexError:
            punkt = ''
        try:
            ter=  grab.doc.select(u'//td[contains(text(),"Район")]/following-sibling::td').text().replace(ray,'')
        except IndexError:
            ter =''
	try:
	    try:
	        uliza = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text().replace(u' [Карта]','')
	    except IndexError:
	        uliza = grab.doc.select(u'//td[contains(text(),"Улица")]/following-sibling::td').text().replace(u' [Карта]','')   
	except IndexError:
	    uliza =''
	try:
	    dom = grab.doc.select(u'//title').text()
        except (IndexError,AttributeError):
	    dom = ''
	    
	try:
            seg = grab.doc.select(u'//dt[contains(text(),"Тип здания:")]/following-sibling::dd[1]').text()
          #print oren
        except DataNotFound:
            seg = '' 
	    
	try:
	    if 'Офисы' in dom:
                naz = grab.doc.select(u'//h2[@class="headtitle"]/a[1]').text()
	    elif 'Гаражи' in dom:
		naz = grab.doc.select(u'//h2[@class="headtitle"]/a[1]').text()
	    else:
		naz = grab.doc.select(u'//h2[@class="headtitle"]/a[2]').text()
          #print naz
        except IndexError:
	    naz = '' 
	    
        try:
            klass = grab.doc.select(u'//dt[contains(text(),"Класс:")]/following-sibling::dd[1]').text()
        except IndexError:
            klass = ''
	    
	try:
	    price = grab.doc.select(u'//td[contains(text(),"Цена:")]/following-sibling::td').text().split(' (')[0]
	  #print price
	except IndexError:
	    price = ''
	    
	try:
            plosh = grab.doc.select(u'//td[contains(text(),"Площадь:")]/following-sibling::td').text()+u' м2'
          #print plosh
        except IndexError:
            plosh = '' 
	    
        try:
            et = grab.doc.select(u'//td[contains(text(),"Этаж / этажей:")]/following-sibling::td').text().split('/')[0]
        except IndexError:
            et = ''
	    
        try:
            et2 = grab.doc.select(u'//td[contains(text(),"Этаж / этажей:")]/following-sibling::td').text().split('/')[1]
        except IndexError:
            et2 = ''
	    
	try:
            opis = grab.doc.select(u'//div[@id="msg_div_msg"]').text()
          #print opis
        except IndexError:
            opis = ''
	    
        try:
            phone = grab.doc.select(u'//span[@id="phone_td_1"]').text().replace('***',str(random.randint(101,999)))
          ##print phone
        except IndexError:
            phone = '' 
	    
        try:
	    try:
	        lico = grab.doc.select(u'//h3[@class="realtor-card__title"][contains(text(),"Представитель: ")]').text().replace(u'Представитель: ','')
	    except IndexError:
	        lico = grab.doc.select(u'//h3[@class="realtor-card__title"]/a[contains(@href,"agents")]').text() 
	except IndexError:
	    lico = ''
	    
	try:
	    try:
		comp = grab.doc.select(u'//h3[@class="realtor-card__title"]/a[contains(@href,"company")]').text()
	    except IndexError:
		comp = grab.doc.select(u'//h4[@class="realtor-card__subtitle"]').text()
	except IndexError:
	    comp = '' 
	try:
	    ohrana =  grab.doc.select(u'//td[contains(text(),"Цена:")]/following-sibling::td').text().split(' (')[1].replace(')','')
	except IndexError:
	    ohrana =''
	try:
	    try:
	        gaz = grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text().replace(u' [Карта]','')
	    except IndexError:
	        gaz = grab.doc.select(u'//td[contains(text(),"Улица")]/following-sibling::td').text().replace(u' [Карта]','')   
	except IndexError:
	    gaz =''
	try:
	    voda =  grab.doc.select(u'//dt[contains(text(),"Состояние:")]/following-sibling::dd[1]').text()
	except IndexError:
	    voda =''
	try:
	    kanal = grab.doc.select(u'//div[@class="cf-object-descr-add"]/span[1]').text()#.split(u'включая')[0]
	except IndexError:
	    kanal =''
	try:
	    elek = grab.doc.select(u'//dt[contains(text(),"Общая площадь:")]/following-sibling::dd[1]').text()
	except IndexError:
	    elek =''
	try:
	    teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
	except IndexError:
	    teplo =''
	    
	try:
            data= grab.doc.select(u'//td[@class="msg_footer"][contains(text(),"Дата:")]').text().split(': ')[1].split(' ')[0]
            #data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(' ','')
		   #print data
        except IndexError:
            data = ''
	    
        try:
            if u'мес' in price:
                oper = u'Аренда'
            else:
	        oper =u'Продажа'
        except IndexError:
            oper = ''
	
	projects = {'url': task.url,
	            'sub': sub,
	            'ray': ray,
	            'punkt': punkt,
	            'teritor': ter,
	            'uliza': uliza,
	            'dom': dom,
	            'seg': seg,
	            'naznachenie': naz,
	            'klass': klass,
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
                    'teplo': teplo,
	            'data':data,
	            'oper':oper
	            
	            }
	yield Task('write',project=projects,grab=grab)
	
    def task_write(self,grab,task):
	
	print('*'*50)
	print  task.project['sub']
	print  task.project['ray']
	print  task.project['punkt']
	print  task.project['teritor']
	print  task.project['uliza']
	print  task.project['dom']
	print  task.project['seg']
	print  task.project['naznachenie']
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
	self.ws.write(self.result, 33, task.project['dom'])
	#self.ws.write(self.result, 6, task.project['orentir'])
	#self.ws.write(self.result, 8, task.project['seg'])
	#self.ws.write(self.result, 8, task.project['tip'])
	self.ws.write(self.result, 9, task.project['naznachenie'])
	#self.ws.write(self.result, 10, task.project['klass'])
	self.ws.write(self.result, 11, task.project['cena'])
	self.ws.write(self.result, 14, task.project['ploshad'])	
	self.ws.write(self.result, 15, task.project['et'])
	self.ws.write(self.result, 16, task.project['ets'])
	#self.ws.write(self.result, 15, task.project['god'])
	#self.ws.write(self.result, 16, task.project['mat'])
	#self.ws.write(self.result, 17, task.project['potolok'])
	#self.ws.write(self.result, 18, task.project['sost'])
	#self.ws.write(self.result, 34, task.project['ohrana'])
	self.ws.write(self.result, 24, task.project['gaz'])
	#self.ws.write(self.result, 18, task.project['voda'])
	#self.ws.write(self.result, 34, task.project['kanaliz'])
	#self.ws.write(self.result, 35, task.project['electr'])
	self.ws.write(self.result, 24, task.project['teplo'])
	self.ws.write_string(self.result, 20, task.project['url'])
	self.ws.write(self.result, 21, task.project['phone'])
	#self.ws.write(self.result, 29, task.project['lico'])
	#self.ws.write(self.result, 30, task.project['company'])
	self.ws.write(self.result, 29, task.project['data'])
	self.ws.write(self.result, 18, task.project['opisanie'])
	self.ws.write(self.result, 19, u'Доска.ру')
	self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	self.ws.write(self.result, 28, task.project['oper'])
	
	
	print('*'*50)
	print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	logger.debug('Tasks - %s' % self.task_queue.size()) 
	print  task.project['oper']
	print('*'*50)
	
	self.result+= 1
	
	
	#if self.result > 100:
	    #self.stop()	
	
	
       
bot = Cian_Com(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
command = 'mount -a'
os.system('echo %s|sudo -S %s' % ('1122', command))
time.sleep(2)
workbook.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/small/sakh_zem.py")
    
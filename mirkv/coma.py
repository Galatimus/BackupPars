#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
import math
import random
import time
import os
from grab import Grab
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)





l= open('links/Com_Arenda.txt').read().splitlines()
oper = u'Аренда'

for i in range(len(l)):
    print '********************************************',i+1,'/',len(l),'*******************************************'
    class MK_Com(Spider):
	def prepare(self):
	    self.f = l[i]
	    for p in range(1,51):
	        try:
                    time.sleep(1)
		    #g = Grab(timeout=10, connect_timeout=20)
		    #g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                    #g.go(self.f)
		    self.sub = self.f.split('/')[3].replace('+',' ')
                    #self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="sort-count"]/span[1]').text())
		    #if int(self.num) > 7000:
			#self.num = 7000
		    #else:
			#self.num = self.num
			
                    #self.pag = int(math.ceil(float(int(self.num))/float(20)))
		    print self.sub
                    #print self.num,self.pag
		    #del g
                    break
                except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError,ValueError):
                    print g.config['proxy'],'Change proxy'
                    print str(p)
		    #del g
                    continue
	    else:
		self.pag = 0
		self.num=0
		self.stop()
	     
	    self.workbook = xlsxwriter.Workbook(u'com/Mirkvartir_Com_'+oper+'_'+str(i+1)+'.xlsx')
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
	    for x in range(100):
                yield Task ('post',url=self.f+'?p=%d'%x,refresh_cache=True,network_try_count=100)
	    yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)
	    
	
	    
	def task_post(self,grab,task):
	    for elem in grab.doc.select(u'//a[@class="offer-title"]'):
		ur = grab.make_url_absolute(elem.attr('href'))
		#print ur
		yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	   
			
      
	    
	    
	    
	    
	def task_item(self, grab, task):

	    
	    
	    try:
	        ray = grab.doc.select(u'//p[@class="address"]/a[contains(text(),"р-н")]').text()
	    except IndexError:
		ray =''
	    try:
		punkt= grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[3]
	    except IndexError:
		punkt = ''
	    try:
		    try:
		        ter= grab.doc.select(u'//a[@class="js-popup-select popup-select InhabitedPoint-popup"]/following::span[@itemprop="name"][1]').text()
		    except IndexError:
		        ter= grab.doc.select(u'//label[contains(text(),"Округ:")]/following-sibling::p/a').text()
	    except IndexError:
		ter =''
		
		
		
	    try:
		uliza = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[4].replace(u'цена','')
	    except IndexError:
		uliza = '' 
	    try:
		dom = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[5].replace(u'цена','')
	    except IndexError:
		dom = ''
		
	    try:
		orentir = grab.doc.select(u'//div[@class="complex-info"]/h3/a[1]').text()
	    except IndexError:
		orentir = ''
		
	    try:
	        seg = grab.doc.select(u'//div[contains(text(),"Инфраструктура")]/following-sibling::div').text()
	    except IndexError:
		seg = '' 
		
	    try:
	        naz = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[2]
	    except IndexError:
		naz = '' 
		
	    try:
		klass = grab.doc.select(u'//div[contains(text(),"ность")]/following-sibling::div[contains(text(),"этаж")]').text()
	    except IndexError:
		klass = ''
		
	    try:
	        price = grab.doc.select(u'//div[@class="price m-all"]').text()
	      #print price
	    except IndexError:
		price = ''
		
	    try:
	        plosh = grab.doc.select(u'//h1').text().split(', ')[0]
	      #print plosh
	    except IndexError:
		plosh = '' 
		
	    try:
		et = grab.doc.select(u'//h1').text().split(', ')[1].split('/')[0]
	    except IndexError:
		et = ''
		
	    try:
	        et2 = grab.doc.select(u'//h1').text().split(', ')[1].split('/')[1].replace(u' этаж','')
	    except IndexError:
		et2 = ''
		
	    try:
	        opis = grab.doc.select(u'//div[@class="l-object-description"]/p').text()
	      #print opis
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
		ohrana = grab.doc.select(u'//p[@class="address"]').text()
	    except IndexError:
		ohrana =''
	    try:
		gaz = grab.doc.select(u'//div[@class="l-object-address"]/p[2]/span/a').text()#.split(', ')[0]
	    except IndexError:
		gaz =''
	    try:
		voda = grab.doc.select(u'//div[@class="l-object-address"]/p[2]/small').text()#.split(', ')[1]
	    except IndexError:
		voda =''
	    try:
		kanal = grab.doc.select(u'//title').text()
	    except IndexError:
		kanal =''
	    try:
		elek = grab.doc.rex_text(u'"lat":(.*?),')
	    except IndexError:
		elek =''
	    try:
		teplo = grab.doc.rex_text(u'"lon":(.*?)}')
	    except IndexError:
		teplo =''
		
	    try:
		try:
		    data = grab.doc.select(u'//div[@class="l-object-dates"]/p[2]').text().split(u' размещено ')[1].split(u' в ')[0]
		except IndexError:
		    data = grab.doc.select(u'//div[@class="dates"]').text().split(u' размещено ')[1].split(u' в ')[0]
	    except IndexError:
		data = ''
	    
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
		        'et': et.replace(klass,''),
	                'ets': et2,
		        'opisanie': opis,
		        'company':comp,
		        'lico':lico.replace(comp,''),
		        'ohrana':ohrana,
	                'body': grab.doc.body,
		        'gaz': gaz,
		        'voda': voda,
		        'kanaliz': kanal,
		        'electr': elek,
		        'teplo': teplo,
		        'data':data.replace('.18','.2018').replace('.19','.2019').replace(u'сегодня', (datetime.today().strftime('%d.%m.%Y'))).replace(u'вчера', '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))}
	    
	    try:
		#url1 = grab.doc.rex_text(u'phoneNumberUrl(.*?)topContact')[4:][:-3]
		#host = re.sub('(?<=ru/).*$','',task.url)
		#ActiveMain = grab.doc.rex_text(u'DetermineActiveMain(.*?);')[2:][:-2]
		#phone_url = host+url1
		#headers ={'Accept': 'application/json, text/plain, */*',
			  #'Accept-Encoding': 'gzip,deflate,br',
			  #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			  #'Cookie': 'activeMain='+ActiveMain,
			  #'Host': host[7:][:-1],
			  #'Referer': task.url,
			  #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0'}
		#gr = Grab()
		#gr.setup(url=phone_url,headers=headers)
		#yield Task('phone',url=task.url+'print/',project=projects,refresh_cache=True,network_try_count=100)
		yield Task('phone',grab=grab,project=projects)
	    except IndexError:
		pass	    
	    
	def task_phone(self, grab, task):
	    #try:
		##phone = re.sub(u'[^\d\+]','',grab.doc.json["phones"][0])
		#phone = re.sub(u'[^\d\+]','',grab.doc.select(u'//td[contains(text(),"Телефон")]/following-sibling::td').text())
		#print 'Phone-OK'
	    #except IndexError:
	    phone = random.choice(list(open('../phone.txt').read().splitlines()))
	    
		
	    yield Task('write',project=task.project,phone=phone,grab=grab)
	    #yield Task('write',url='https://mini.s-shot.ru/T90/FS0/1024x0/JPEG/1024/Z100/?'+task.project['url'],project=task.project,phone=phone,refresh_cache=True,network_try_count=100)
	    
	def task_write(self,grab,task):
	    
	    print('*'*50)
	    print  task.project['sub']
	    print  task.project['ray']
	    print  task.project['punkt']
	    #print  task.project['teritor']
	    print  task.project['uliza']
	    print  task.project['dom']
	    print  task.project['orentir']
	    print  task.project['seg']
	    print  task.project['naznachenie']
	    #print  task.project['klass']
	    print  task.project['cena']
	    print  task.project['ploshad']
	    print  task.project['et']
	    print  task.project['ets']
	    print  task.project['opisanie']
	    print  task.project['url']
	    print  task.phone
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
	    #self.ws.write(self.result, 3, task.project['teritor'])
	    self.ws.write(self.result, 4, task.project['uliza'])
	    self.ws.write(self.result, 5, task.project['dom'])
	    self.ws.write(self.result, 6, task.project['orentir'])
	    self.ws.write(self.result, 8, task.project['seg'])
	    #self.ws.write(self.result, 8, task.project['tip'])
	    self.ws.write(self.result, 9, task.project['naznachenie'])
	    self.ws.write(self.result, 16, task.project['klass'])
	    self.ws.write(self.result, 11, task.project['cena'])
	    self.ws.write(self.result, 14, task.project['ploshad'])	
	    self.ws.write(self.result, 15, task.project['et'])
	    self.ws.write(self.result, 16, task.project['ets'])
	    #self.ws.write(self.result, 15, task.project['god'])
	    #self.ws.write(self.result, 16, task.project['mat'])
	    #self.ws.write(self.result, 17, task.project['potolok'])
	    #self.ws.write(self.result, 18, task.project['sost'])
	    self.ws.write(self.result, 24, task.project['ohrana'])
	    self.ws.write(self.result, 26, task.project['gaz'])
	    self.ws.write(self.result, 27, task.project['voda'])
	    self.ws.write(self.result, 33, task.project['kanaliz'])
	    self.ws.write(self.result, 34, task.project['electr'])
	    self.ws.write(self.result, 35, task.project['teplo'])
	    self.ws.write_string(self.result, 20, task.project['url'])
	    self.ws.write(self.result, 21, task.phone)
	    self.ws.write(self.result, 22, task.project['lico'])
	    self.ws.write(self.result, 23, task.project['company'])
	    self.ws.write(self.result, 29, task.project['data'])
	    self.ws.write(self.result, 18, task.project['opisanie'])
	    self.ws.write(self.result, 19, u'MIRKVARTIR.RU')
	    self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	    self.ws.write(self.result, 28, oper)
	    
	    
	    print('*'*50)
            print 'Ready - '+str(self.result)#+'/'+str(self.num)
            logger.debug('Tasks - %s' % self.task_queue.size()) 
            print '***',i+1,'/',len(l),'***'
            print oper
            print('*'*50)
	    
	    self.result+= 1
	    
	    #if str(self.result) == str(self.num):
		#self.stop()

	   
    bot = MK_Com(thread_number=5,network_try_limit=1000)
    bot.load_proxylist('../tipa.txt','text_file')
    bot.create_grab_instance(timeout=500, connect_timeout=500)
    try:
        bot.run()
    except KeyboardInterrupt:
        pass
    print('Wait 2 sec...')
    time.sleep(1)
    print('Save it...')    
    #command = 'mount -a'
    #os.system('echo %s|sudo -S %s' % ('1122', command))
    time.sleep(2)
    bot.workbook.close()
    print('Done')    

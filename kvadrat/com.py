#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
import os
from datetime import datetime
from grab import Grab
import xlsxwriter
import time
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


i = 0
l= ['http://kvadrat22.ru/mob_sellcombank-1000-1.html',
    'http://kvadrat24.ru/mob_sellcombank-1000-1.html',
    'http://kvadrat54.ru/mob_sellcombank-1000-1.html',
    'http://kvadrat64.ru/mob_sellcombank-1000-1.html',
    'http://kvadrat66.ru/mob_sellcombank-1000-1.html',
    'http://kvadrat72.ru/mob_sellcombank-1000-1.html',
    'http://kvadrat74.ru/mob_sellcombank-1000-1.html',
    'http://n30.ru/mob_sellcombank-1000-1.html',
    'http://kemdom.ru/mob_sellcombank-1000-1.html',
    'http://n002.ru/mob_sellcombank-1000-1.html',
    'http://kazan-n.ru/mob_sellcombank-1000-1.html',
    'http://nd27.ru/mob_sellcombank-1000-1.html',
    #'http://nd23.ru/mob_sellcombank-1000-1.html',
    'http://kvadrat22.ru/mob_givecombank-1000-1.html',
    'http://kvadrat24.ru/mob_givecombank-1000-1.html',
    'http://kvadrat54.ru/mob_givecombank-1000-1.html',
    'http://kvadrat64.ru/mob_givecombank-1000-1.html',
    'http://kvadrat66.ru/mob_givecombank-1000-1.html',
    'http://kvadrat72.ru/mob_givecombank-1000-1.html',
    'http://kvadrat74.ru/mob_givecombank-1000-1.html',
    'http://n30.ru/mob_givecombank-1000-1.html',
    #'http://nd23.ru/mob_givecombank-1000-1.html',
    'http://kemdom.ru/mob_givecombank-1000-1.html',
    'http://n002.ru/mob_givecombank-1000-1.html',
    'http://kazan-n.ru/mob_givecombank-1000-1.html',
    'http://nd27.ru/mob_givecombank-1000-1.html',
    'http://n30.ru/mob_givecombank-1000-1.html']

page = l[i]

while True:
    print '********************************************',i+1,'/',len(l),'*******************************************'
    class Kvadrat_Com(Spider):
	def prepare(self):
	    self.f = page
	    for p in range(1,51):
		try:
		    time.sleep(1)
		    g = Grab(timeout=20, connect_timeout=20)
		    g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
		    g.go(self.f)
		    conv = [(u'Хабаровска',u'Хабаровский край'),(u'Барнаула',u'Алтайский край'),
			    (u'Красноярска',u'Красноярский край'),(u'Саратова',u'Саратовская область'),
			    (u'Новосибирска',u'Новосибирская область'),(u'Екатеринбурга',u'Свердловская область'),
			    (u'Тюмени',u'Тюменская область'),(u'Челябинска',u'Челябинская область'),
			    (u'Астрахани',u'Астраханская область'),(u'Кемерово',u'Кемеровская область'),
			    (u'Уфы',u'Башкортостан'),(u'Казани',u'Татарстан'),(u'Краснодара',u'Краснодарский край')]        
		    dt = g.doc.select(u'//a[@class="amenu"][contains(text(),"Недвижимость")]').text().replace(u'Недвижимость','') 
		    self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
		    print self.sub
		    del g
		    break
		except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
		    print g.config['proxy'],'Change proxy'
		    g.change_proxy()
		    del g
		    continue
	    else:
		self.sub = ''
		
	    self.workbook = xlsxwriter.Workbook(u'com/Kvadrat_%s' % bot.sub +str(i+1)+ u'_Коммерческая.xlsx')
	    self.ws = self.workbook.add_worksheet(u'Kvadrat_Коммерческая')
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
	    yield Task ('post',url = page,network_try_count=100)
	    
	    
	    
	def task_page(self,grab,task):
	    try:
		pg = grab.doc.select(u'//div[@class="dphase"]/following-sibling::a[1]')
		u = grab.make_url_absolute(pg.attr('href'))
		yield Task ('post',url= u,network_try_count=100)
	    except DataNotFound:
		print('*'*100)
		print '!!!','NO PAGE NEXT','!!!'
		print('*'*100)
		logger.debug('%s taskq size' % self.task_queue.size())	
	    
	def task_post(self,grab,task):
	    for elem in grab.doc.select(u'//a[@class="site3"]'):
		ur = grab.make_url_absolute(elem.attr('href'))  
		#print ur
		yield Task('item', url=ur,network_try_count=100)
	    yield Task("page", grab=grab,network_try_count=100,use_proxylist=False)
			
      
	    
	    
	    
	    
	def task_item(self, grab, task):
	    
	   
	    try:
		orent = grab.doc.select(u'//a[@class="blue"][contains(@href, "combank")]').text()
	    except IndexError:
		orent = ''
	    
	    try:
	        seg = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Объект: ')[1].split(' (')[0].split(u' Этажи')[0]
	      #print oren
	    except IndexError:
		seg = '' 
		
	    try:
	        naz = grab.doc.select(u'//div[@class="divdec"]/div[contains(text(),"Назначение:")]').text().split(u'Назначение: ')[1]
	      #print naz
	    except IndexError:
		naz = '' 
		
	    try:
		klass = grab.doc.select(u'//div[@class="divdec"]/div/span[contains(text(),"класс")]').text().split(' (')[1].replace(')','')
	    except IndexError:
		klass = ''
		
	    try:
	        price = grab.doc.select(u'//td[@class="thprice"]').text()
	      #print price
	    except IndexError:
		price = ''
		
	    try:
		try:
	            plosh = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Продаваемая площадь: ')[1].split(u' м²')[0]+u' м2'
	        except IndexError:
		    plosh = grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Сдаваемая площадь: ')[1].split(u' м²')[0]+u' м2'
	    except IndexError:
		plosh = '' 
		
	    try:
		et = re.sub('[^\d\/]','',grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Этажи/этажей в строении: ')[1]).split('/')[0][:1]
	    except IndexError:
		et = ''
		
	    try:
		mat = re.sub('[^\d\/]','',grab.doc.select(u'//div[@class="divdec"]/div').text().split(u'Этажи/этажей в строении: ')[1]).split('/')[1]
	    except IndexError:
		mat = ''
		
	    try:
	        opis = grab.doc.select(u'//div[contains(text(), "Дополнительная информация:")]/span').text()
	      #print opis
	    except IndexError:
		opis = ''
		
	    try:
		tip = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[0])
		user = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[1])
		pkey = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[2])
		link = task.url.split(u'mob')[0]#+u'ru'
		url_ph = link+'showphone.php?tip='+tip+'&id='+user+'&from='+pkey
		g2 = grab.clone()
		g2.go(url_ph)
	        phone = re.sub('[^\d\,]','',re.findall('innerHTML="(.*?)"',g2.doc.body)[0])
		del g2
	      ##print phone
	    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
		phone = '' 
		
	    try:
		try:
		    lico = grab.doc.select(u'//div[@class="divdec"]').text().split(u'Персона для контактов:')[1].split(u'Контактный')[0]
		except IndexError:
		    lico = grab.doc.rex_text(u'Персона для контактов(.*?)/').split('=d>')[1].replace('<','')
	    except IndexError:
		lico = ''
		
	    try:
	        gaz = grab.doc.select(u'//td[@class="hh"]').text().split(u'), ')[1].replace(u' на карте','')
	    except IndexError:
		gaz =''
	    try:
		#voda =  re.sub(u'[^\d\-]','',grab.doc.select(u'//td[@class="tdate"]').text().split(u'создано ')[1]).replace('-','.')
		voda =  grab.doc.rex_text(u'создано (.*?)</td>').replace('-','.')
	    except IndexError:
		voda =''
	    try:
		kanal = grab.doc.select(u'//td[@class="hh"]/text()').text()
	    except IndexError:
		kanal =''
	    try:
		elek = grab.doc.select(u'//div[@class="divdec"]/div[contains(text(),"Парковка:")]').text().split(u'Парковка: ')[1]
	    except IndexError:
		elek =''
	    try:
		teplo = grab.doc.select(u'//div[@class="divdec"]/div[contains(text(),"Агенство недвижимости:")]').text().split(u'Агенство недвижимости: ')[1]
	    except IndexError:
		teplo =''
		
	    try:
		#data = re.sub(u'[^\d\-]','',grab.doc.select(u'//td[@class="tdate"]').text().split(u'обновлено ')[0]).replace('-','.')
		data = grab.doc.rex_text(u'обновлено (.*?)<br>').replace('-','.')
	    except IndexError:
		data = ''
		
	    try:
		oper = grab.doc.select(u'//div[@class="a"]').text().split(' ')[0]
	    except IndexError:
		oper = ''
	    try:
	        istoch = grab.doc.select(u'//a[@class="amenu"][contains(text(),"Недвижимость")]').text().replace(u'Недвижимость',u'Недвижимость ')
	    except IndexError:
		istoch = ''		
	    
	    projects = {'url': task.url,
		        'sub': self.sub,
		        'orentir': orent,
		        'seg': seg,
		        'naznachenie': naz,
		        'klass': klass,
		        'cena': price,
		        'ploshad': plosh,
		        'et': et,
		        'mat': mat,
		        'opisanie': opis,
	                'istochnik': istoch,
		        'phone':phone,
		        'lico':lico,
		        'gaz': gaz,
		        'voda': voda,
		        'kanaliz': kanal,
		        'electr': elek,
		        'teplo': teplo,
		        'data':data,
		        'oper':oper}
	    
	    #try:
		#tip = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[0])
		#user = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[1])
		#pkey = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[2])
		#link = task.url.split(u'mob')[0]#+u'ru'
		#url_ph = link+'showphone.php?tip='+tip+'&id='+user+'&from='+pkey
		#headers ={'Accept': '*/*',
			  #'Accept-Encoding': 'gzip, deflate',
			  #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			  ##'Cookie': 'PHPSESSID='+key+'.'+pkey,
			  #'Host': link,
			  #'Referer': task.url,
			  #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0'}		    
		#gr = Grab()
		#gr.setup(url=url_ph,headers=headers)	            
		#yield Task('phone',grab=gr,refresh_cache=True,network_try_count=10)
	    #except IndexError:
	        #yield Task('phone',grab=grab)	    
	    
	    
	    try:
		    ad= grab.doc.select(u'//td[@class="hh"]/text()').text().split('), ')[1].replace(u' на карте','')
		    link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ad
		    yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=100)
	    except IndexError:
		yield Task('adres',grab=grab,project=projects)
		    
		    
	#def task_phone(self, grab, task):
	    #try:
		##phone =  re.sub('[^\d]','',re.findall("tel:(.*?)>",grab.response.body)[0])
		#self.phone = grab.response.body
		#print self.phone
	    #except IndexError:
		#self.phone = ''
			
			
	def task_adres(self, grab, task):
	    try:   
		punkt= grab.doc.rex_text(u'LocalityName":"(.*?)"')
	    except IndexError:
		punkt = ''
	    try:
		ter=  grab.doc.rex_text(u'DependentLocalityName":"(.*?)"')
	    except IndexError:
		ter =''
	    try:
	        uliza=grab.doc.rex_text(u'ThoroughfareName":"(.*?)"')
	    except IndexError:
		uliza = ''
	    try:
		dom = grab.doc.rex_text(u'PremiseNumber":"(.*?)"')
	    except IndexError:
		dom = ''
		
	    project2 ={'punkt':punkt,
	                'teritor': ter,
	               'ulica':uliza,
	               'dom':dom}		    
	    
	    
		    
	    yield Task('write',project=task.project,proj=project2,grab=grab)
	    
	def task_write(self,grab,task):
	    
	    print('*'*50)
	    print  task.project['sub']
	    print  task.proj['punkt']
	    print  task.proj['teritor']
	    print  task.proj['ulica']
	    print  task.project['orentir']
	    print  task.proj['dom']
	    print  task.project['seg']
	    print  task.project['naznachenie']
	    print  task.project['klass']
	    print  task.project['cena']
	    print  task.project['ploshad']
	    print  task.project['et']
	    print  task.project['mat']
	    print  task.project['opisanie']
	    print  task.project['url']
	    print  task.project['phone']
	    print  task.project['lico']
	    print  task.project['kanaliz']
	    print  task.project['electr']
	    print  task.project['teplo']
	    print  task.project['data']
	    print  task.project['gaz']
	    print  task.project['voda']	    
	    print  task.project['istochnik']
	    #print self.phone
	    
	    
	    
	    self.ws.write(self.result, 0, task.project['sub'])
	    #self.ws.write(self.result, 1, task.project['ray'])
	    self.ws.write(self.result, 4, task.proj['ulica'])
	    self.ws.write(self.result, 3, task.proj['teritor'])
	    self.ws.write(self.result, 2, task.proj['punkt'])
	    self.ws.write(self.result, 5, task.proj['dom'])
	    self.ws.write(self.result, 7, task.project['orentir'])
	    self.ws.write(self.result, 8, task.project['seg'])
	    #self.ws.write(self.result, 8, task.project['tip'])
	    self.ws.write(self.result, 9, task.project['naznachenie'])
	    self.ws.write(self.result, 10, task.project['klass'])
	    self.ws.write(self.result, 11, task.project['cena'])
	    self.ws.write(self.result, 14, task.project['ploshad'])	
	    self.ws.write(self.result, 15, task.project['et'])
	    self.ws.write(self.result, 16, task.project['mat'])
	    #self.ws.write(self.result, 15, task.project['god'])
	    #self.ws.write(self.result, 16, task.project['mat'])
	    #self.ws.write(self.result, 17, task.project['potolok'])
	    #self.ws.write(self.result, 18, task.project['sost'])
	    #self.ws.write(self.result, 19, task.project['ohrana'])
	    self.ws.write(self.result, 24, task.project['gaz'])
	    self.ws.write(self.result, 29, task.project['voda'])
	    self.ws.write(self.result, 33, task.project['kanaliz'])
	    self.ws.write(self.result, 37, task.project['electr'])
	    self.ws.write(self.result, 23, task.project['teplo'])
	    self.ws.write_string(self.result, 20, task.project['url'])
	    self.ws.write(self.result, 21, task.project['phone'])
	    self.ws.write(self.result, 22, task.project['lico'])
	    #self.ws.write(self.result, 30, task.project['company'])
	    self.ws.write(self.result, 30, task.project['data'])
	    self.ws.write(self.result, 18, task.project['opisanie'])
	    self.ws.write(self.result, 19, task.project['istochnik'])
	    self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	    self.ws.write(self.result, 28, task.project['oper'])
	    
	    print('*'*50)
	    print 'Ready - '+str(self.result)
	    logger.debug('Tasks - %s' % self.task_queue.size()) 
	    print '***',i+1,'/',len(l),'***'
	    print  task.project['oper']
	    print('*'*100)
	    self.result+= 1
	
	    #if self.result > 7:
		#self.stop()

	   
    bot = Kvadrat_Com(thread_number=5,network_try_limit=1000)
    bot.load_proxylist('../tipa.txt','text_file')
    bot.create_grab_instance(timeout=50, connect_timeout=50)
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
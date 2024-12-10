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
import random
import time
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



#l= ['https://www.beboss.ru/kn/ru_office_sell',
    #'https://www.beboss.ru/kn/ru_retail_sell',
    #'https://www.beboss.ru/kn/ru_stock_sell',
    #'https://www.beboss.ru/kn/ru_industry_sell',
    #'https://www.beboss.ru/kn/ru_spec_sell',
    #'https://www.beboss.ru/kn/ru_office_rent',
    #'https://www.beboss.ru/kn/ru_retail_rent',
    #'https://www.beboss.ru/kn/ru_stock_rent',
    #'https://www.beboss.ru/kn/ru_industry_rent',
    #'https://www.beboss.ru/kn/ru_spec_rent']
    

i = 0

l= ['tlt','samara','krym','vladimir-obl','vladivostok','msk','tmn','kemerovo-obl']





dc = len(l)
page = l[i]

while True:
    print '********************************************',i+1,'/',len(l),'*******************************************'
    class Kvadrat_Com(Spider):
	def prepare(self):
	    self.f = page
	    #for p in range(1,51):
		#try:
		    #time.sleep(1)
		    #g = Grab(timeout=20, connect_timeout=20)
		    #g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
		    #g.go(self.f)
		    ##self.oper = g.doc.select(u'//h1').text().split(' ')[0]
		    #self.seg = g.doc.select(u'//h1').text().replace(self.oper,'').split(u' в ')[0]
		    #print self.oper,self.seg
		    #del g
		    #break
		#except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
		    #print g.config['proxy'],'Change proxy'
		    #g.change_proxy()
		    #del g
		    #continue
	    #else:
		#self.seg = ''
		
	    self.workbook = xlsxwriter.Workbook(u'com/Beboss_Comm_'+str(i+1)+'.xlsx')
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
	    self.ws.write(0, 36, u"ТРАССА")
	    self.ws.write(0, 37, u"ПАРКОВКА")
	    self.ws.write(0, 38, u"ОХРАНА")
	    self.ws.write(0, 39, u"КОНДИЦИОНИРОВАНИЕ")
	    self.ws.write(0, 40, u"ИНТЕРНЕТ")
	    self.ws.write(0, 41, u"ТЕЛЕФОН (КОЛИЧЕСТВО ЛИНИЙ)")
	    self.ws.write(0, 42, u"УСЛУГИ")
	    self.ws.write(0, 43, u"НАЛИЧИЕ ОТДЕЛКИ ПОМЕЩЕНИЙ")
	    self.ws.write(0, 44, u"ОТДЕЛЬНЫЙ ВХОД")
	    self.ws.write(0, 45, u"ВЫСОТА ПОТОЛКОВ")	    
	    self.result= 1
	    
		
		
		
		  
	
	def task_generator(self):
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_office_sell',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_retail_sell',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_stock_sell',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_industry_sell',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_spec_sell',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_office_rent',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_retail_rent',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_stock_rent',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_industry_rent',refresh_cache=True,network_try_count=100)
	    yield Task ('post',url = 'https://www.beboss.ru/kn/'+page+'_spec_rent',refresh_cache=True,network_try_count=100)
	    
	def task_post(self,grab,task):
	    for elem in grab.doc.select(u'//div[@class="obj__right"]/a[contains(text(),"Подробнее")]'):
		ur = grab.make_url_absolute(elem.attr('href'))  
		#print ur
		yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	    yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	    
	    
	    
	def task_page(self,grab,task):
	    try:
		pg = grab.doc.select(u'//span[contains(text(),"Следующая")]/ancestor::a')
		u = grab.make_url_absolute(pg.attr('href'))
		yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	    except IndexError:
		print('*'*10)
		print '!!!','NO PAGE NEXT','!!!'
		print('*'*10)
		logger.debug('%s taskq size' % self.task_queue.size())	

	    
	def task_item(self, grab, task):
	    
	   
	    try:
		orent = grab.doc.select(u'//p[@class="object-addresses"]').text()
	    except IndexError:
		orent = ''
	    
	    try:
	        seg = grab.doc.select(u'//dt[contains(text(),"Класс здания")]/following-sibling::dd[1]').text()
	    except IndexError:
		seg = '' 
		
	    try:
	        naz = grab.doc.select(u'//p[@class="kn-obj-info b-franchise-hide-mobile"]/text()[1]').text().split(': ')[1]
	      #print naz
	    except IndexError:
		naz = '' 
		
	    try:
		klass = grab.doc.select(u'//p[@class="kn-obj-title b-franchise-hide-mobile"]').text().split(': ')[1]
	    except IndexError:
		klass = ''
		
	    try:
	        price = grab.doc.select(u'//dt[contains(text(),"Этаж расположения помещения")]/following-sibling::dd[1]').text()
	      #print price
	    except IndexError:
		price = ''
		
	    try:
                plosh = grab.doc.select(u'//dt[contains(text(),"Введено в эксплуатацию в")]/following-sibling::dd[1]').text()
	    except IndexError:
		plosh = '' 
		
	    try:
		et = grab.doc.select(u'//p[@itemprop="description"]').text()
	    except IndexError:
		et = ''
		
	    try:
		mat = grab.doc.select(u'//p[@class="franchise-person__name"]').text()
	    except IndexError:
		mat = ''
		
	    try:
	        opis = grab.doc.select(u'//div[@class="block"]/p[2]').text().split(' (')[0]
	    except IndexError:
		opis = ''

	    try:
	        lico = grab.doc.select(u'//div[@class="block"]/p[2]').text().split(' (')[1].replace(')','')
	    except IndexError:
		lico = ''
		
	    try:
	        gaz = grab.doc.select(u'//span[@class="kn-type-object__date"]').text().split(', ')[0].replace(u'Обновлено ','')
	    except IndexError:
		gaz =''
	    try:
		voda =  grab.doc.select(u'//h1').text()
	    except IndexError:
		voda =''
	    try:
		kanal = grab.doc.rex_text(u'id="lat" value="(.*?)"')
	    except IndexError:
		kanal =''
	    try:
		elek = grab.doc.rex_text(u'id="lng" value="(.*?)"')
	    except IndexError:
		elek =''
	    try:
		teplo = grab.doc.select(u'//dt[contains(text(),"Имеется парковка")]/following-sibling::dd[1]').text()
	    except IndexError:
		teplo =''
		
	    try:
		try:
	            oper = grab.doc.select(u'//div[@class="kn-type-object"]/span[1]').text()
		except IndexError:
		    oper = grab.doc.select(u'//span[@class="kn-type-new"]').text()
	    except IndexError:
		oper =''
		
	    try:
	        tip =  grab.doc.select(u'//title').text().split(', ')[0]
	    except IndexError:
		tip =''	    
		
	    try:
		data = grab.doc.select(u'//dt[contains(text(),"Безопасность")]/following-sibling::dd[1]').text()
	    except IndexError:
		data = ''
		
	    try:
		cond = grab.doc.select(u'//dt[contains(text(),"Вентиляция и кондиционирование")]/following-sibling::dd[1]').text()
	    except IndexError:
		cond = ''
	    try:
	        istoch = grab.doc.select(u'//dt[contains(text(),"Ремонт помещения")]/following-sibling::dd[1]').text()
	    except IndexError:
		istoch = ''
		
	    if 'ID' in gaz:
		gaz = ''
	    else:
		gaz=gaz
		
	    if 'ID' in oper:
		oper = u'Аренда'
	    else:
		oper=oper	    
	    
	    projects = {'url': task.url,
		        'orentir': orent,
		        'seg': seg,
		        'naznachenie': naz,
		        'klass': klass,
		        'cena': price,
		        'ploshad': plosh,
		        'et': et,
	                'condey': cond,
		        'mat': mat,
		        'opisanie': opis,
	                'istochnik': istoch,
		        'phone': random.choice(list(open('../phone.txt').read().splitlines())),
		        'lico':lico,
		        'gaz': gaz,
		        'voda': voda,
		        'kanaliz': kanal,
		        'electr': elek,
		        'teplo': teplo,
	                'segment': tip.replace(oper,''),
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
		    link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+orent
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
	        sub= grab.doc.rex_text(u'AdministrativeAreaName":"(.*?)"')
	    except IndexError:
	        sub = ''	    
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
	               'sub': sub,
	                'teritor': ter,
	               'ulica':uliza,
	               'dom':dom}		    
	    
	    
		    
	    yield Task('write',project=task.project,proj=project2,grab=grab)
	    
	def task_write(self,grab,task):
	    
	    print('*'*50)
	    print  task.proj['sub']
	    print  task.proj['punkt']
	    print  task.proj['teritor']
	    print  task.proj['ulica']	    
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
	    print  task.project['orentir']
	    print  task.project['segment']
	    #print self.phone
	    
	    
	    
	    self.ws.write(self.result, 0, task.proj['sub'])
	    #self.ws.write(self.result, 1, task.project['ray'])
	    self.ws.write(self.result, 4, task.proj['ulica'])
	    self.ws.write(self.result, 3, task.proj['teritor'])
	    self.ws.write(self.result, 2, task.proj['punkt'])
	    self.ws.write(self.result, 5, task.proj['dom'])
	    self.ws.write(self.result, 24, task.project['orentir'])
	    self.ws.write(self.result, 10, task.project['seg'])
	    self.ws.write(self.result, 7, task.project['segment'])
	    self.ws.write(self.result, 11, task.project['naznachenie'])
	    self.ws.write(self.result, 14, task.project['klass'])
	    self.ws.write(self.result, 15, task.project['cena'])
	    self.ws.write(self.result, 17, task.project['ploshad'])	
	    self.ws.write(self.result, 18, task.project['et'])
	    self.ws.write(self.result, 22, task.project['mat'])
	    self.ws.write(self.result, 39, task.project['condey'])
	    self.ws.write(self.result, 43, task.project['istochnik'])
	    #self.ws.write(self.result, 17, task.project['potolok'])
	    #self.ws.write(self.result, 18, task.project['sost'])
	    #self.ws.write(self.result, 19, task.project['ohrana'])
	    self.ws.write(self.result, 30, task.project['gaz'])
	    self.ws.write(self.result, 33, task.project['voda'])
	    self.ws.write(self.result, 34, task.project['kanaliz'])
	    self.ws.write(self.result, 35, task.project['electr'])
	    self.ws.write(self.result, 37, task.project['teplo'])
	    self.ws.write_string(self.result, 20, task.project['url'])
	    self.ws.write(self.result, 21, task.project['phone'])
	    self.ws.write(self.result, 27, task.project['lico'])
	    #self.ws.write(self.result, 30, task.project['company'])
	    self.ws.write(self.result, 38, task.project['data'])
	    self.ws.write(self.result, 26, task.project['opisanie'])
	    self.ws.write(self.result, 19, 'БИБОСС')
	    self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	    self.ws.write(self.result, 28, task.project['oper'])
	    
	    print('*'*50)
	    print 'Ready - '+str(self.result)
	    logger.debug('Tasks - %s' % self.task_queue.size()) 
	    print '***',i+1,'/',len(l),'***'
	    print  task.project['oper']
	    print('*'*100)
	    self.result+= 1
	
	    #if self.result > 20:
		#self.stop()

	   
    bot = Kvadrat_Com(thread_number=5,network_try_limit=2000)
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
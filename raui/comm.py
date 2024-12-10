#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
import time
import random
import xlsxwriter
from datetime import datetime,timedelta
import sys
import os
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


i = 0
l= open('links/comm.txt').read().splitlines()

page = l[i] 




while True:
     print '*******************',i+1,'/',len(l),'*****************************'
     class Raui_Com(Spider):
	  def prepare(self):
	       self.urls = page+'?per_page=50'
	       if 'kupit' in self.urls:
		    self.oper = u'Продажа' 
	       elif 'snyat' in self.urls:
		    self.oper = u'Аренда'
	       print self.oper
	       self.workbook = xlsxwriter.Workbook(u'com/Raui_'+self.oper+str(i+1)+'.xlsx') 
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
	       self.ws.write(0, 44, u"СИСТЕМА ВЕНТИЛЯЦИИ")
	       self.ws.write(0, 45, u"ОТДЕЛЬНЫЙ ВХОД")
	       self.ws.write(0, 46, u"ВИД СОБСТВЕННОСТИ")
	       self.ws.write(0, 47, u"ЛИФТ")
	       self.ws.write(0, 48, u"ВЫСОТА ПОТОЛКОВ")  
	       #self.r = conv     
		    
	       self.result= 1
	     
		    
	 
	  def task_generator(self):
	       #for x in range(1,int(self.num)+1):
	       yield Task ('post',url=self.urls,refresh_cache=True,network_try_count=100)
       
		    
				   
	  def task_post(self,grab,task):    
	       for elem in grab.doc.select(u'//a[contains(text(),"Подробнее")]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	       
	       
	  def task_page(self,grab,task):
	       try:         
		    pg = grab.doc.select(u'//a[@class="pagging__link"]/span[contains(text(),"»")]/ancestor::a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print('*'*50)
		    print '!!!','NO PAGE NEXT','!!!'
		    print('*'*50)	  
     
	  def task_item(self, grab, task):
	      
	       try:
		    sub= grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[2]').text()
	       except IndexError:
		    sub =''	 
	       try:
		    try:
			 try:
			      ray = grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[contains(text()," район")]').text()
			 except IndexError:
			      ray = grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[contains(text()," округ ")]').text()
		    except IndexError:
			 ray = grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[contains(text()," образование ")]').text()
	       except IndexError:
		    ray = ''          
	       try:
		    if sub == u'Москва':
			 punkt= u'Москва'
		    elif sub == u'Санкт-Петербург':
			 punkt= u'Санкт-Петербург'
		    elif sub == u'Севастополь':
			 punkt= u'Севастополь'
		    else:
			 punkt= grab.doc.select(u'//div[@class="breadcrumbs"]/div/a[4]').text()
	       except IndexError:
		    punkt = ''
		    
	       try:
		    ter= grab.doc.select(u'//td[contains(text(),"Тип строения:")]/following-sibling::td').text()
	       except IndexError:
		    ter =''
		    
	       try:
		    uliza = grab.doc.select(u'//td[contains(text(),"Класс строения:")]/following-sibling::td').text()
	       except IndexError:
		    uliza = ''
		    
	       try:
		    dom = grab.doc.select(u'//div[@class="item__priceinfo"]').text()
	       except IndexError:
		    dom = ''
		    
	       try:
		    trassa = grab.doc.select(u'//td[contains(text(),"Этаж:")]/following-sibling::td').text().split(' / ')[0]
		     #print rayon
	       except IndexError:
		    trassa = ''
		    
	       try:
		    udal = grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(u' купить,')[0]
	       except IndexError:
		    udal = ''
	       try:
		    seg = grab.doc.select(u'//td[contains(text(),"Этаж:")]/following-sibling::td').text().split(' / ')[1]
	       except IndexError:
		    seg = ''	       
		    
	       try:
		    price = grab.doc.select(u'//div[@class="item__price"]').text()#.replace(u'a',u'р.')
	       except IndexError:
		    price = ''
		    
	       try:
		    plosh = grab.doc.select(u'//td[contains(text(),"Площадь:")]/following-sibling::td').text()
	       except IndexError:
		    plosh = '' 
	       try:
		    cena_za = grab.doc.select(u'//td[contains(text(),"Год постройки:")]/following-sibling::td').text()
	       except IndexError:
		    cena_za = '' 
		    
	       
	       try:
		    ohrana = grab.doc.select(u'//p[contains(text(),"Агентство")]/following-sibling::h3[1]').text()
	       except IndexError:
		    ohrana =''
	       try:
		    gaz = grab.doc.select(u'//p[contains(text(),"Имя:")]/following-sibling::h3[1]').text()
	       except IndexError:
		    gaz =''
	       try:
		    voda = grab.doc.select(u'//div[@class="item__metro"]').text().split(', ')[0]
	       except IndexError:
		    voda =''
	       try:
		    kanal = grab.doc.select(u'//div[@class="item__metro"]').text().split(', ')[1]
	       except IndexError:
		    kanal =''
	       try:
		    elek = grab.doc.select(u'//title').text()
	       except IndexError:
		    elek =''
		    
	       try:
		    lat = grab.doc.select(u'//div[@id="map_item"]').attr('data-croods').split(',')[0]
	       except IndexError:
		    lat =''
	  
	       try:
		    lng = grab.doc.select(u'//div[@id="map_item"]').attr('data-croods').split(',')[1]
	       except IndexError:
		    lng =''	 
		    
		    
	       try:
		    teplo = grab.doc.select(u'//h1').text()
	       except IndexError:
		    teplo =''  
			 
	       try:
		    opis = grab.doc.select(u'//div[@class="item-text"]').text() 
	       except IndexError:
		    opis = ''
		    
	      
		    
	       try:
		    data= datetime.strptime(grab.doc.select(u'//meta[@property="article:published_time"]').attr('content')[:10].replace('-','.'), '%Y.%m.%d')
	       except IndexError:
		    data = ''
		    
	       try:
		    park = grab.doc.select(u'//td[contains(text(),"Парковка:")]/following-sibling::td').text() 
	       except IndexError:
		    park = '' 
		    
	       try:
		    usl = grab.doc.select(u'//td[contains(text(),"Охрана:")]/following-sibling::td').text() 
	       except IndexError:
		    usl = ''	 
		    
	       try:
		    inet = grab.doc.select(u'//td[contains(text(),"Интернет:")]/following-sibling::td').text()
	       except IndexError:
		    inet =''	
		    
	       try:
		    lini = grab.doc.select(u'//td[contains(text(),"Кол-во тел.линий:")]/following-sibling::td').text()#.split(u' из ')[1]
	       except IndexError:
		    lini =''
		    
		    
	       try:
		    otd = grab.doc.select(u'//td[contains(text(),"Состояние:")]/following-sibling::td').text() 
	       except IndexError:
		    otd = ''
		    
	       try:
		    vxod = grab.doc.select(u'//td[contains(text(),"Вход:")]/following-sibling::td').text()#.split(u' из ')[1]
	       except IndexError:
		    vxod =''
		    
	       try:
		    lift = grab.doc.select(u'//td[contains(text(),"Лифт:")]/following-sibling::td').text()#.split(u' из ')[1]
	       except IndexError:
		    lift =''	
		    
	       try:
		    pot = grab.doc.select(u'//td[contains(text(),"Высота потолков:")]/following-sibling::td').text()#.split(u' из ')[1]
	       except IndexError:
		    pot =''		  
		    
		    
	       try:
		    phone = re.sub('[^\d]','',grab.doc.select(u'//div[@class="item-contacts-values"]/div').text()+str(random.randint(100000,999999)))
	       except IndexError:
		    phone = random.choice(list(open('../phone.txt').read().splitlines()))	  
		    
		    
	       #id_phone = re.sub('[^\d]','',grab.doc.select(u'//div[@class=" contactstabs"]').attr('data-id'))
	       
	       #phone_url = 'https://raui.ru/ajax/item/contact?id='+id_phone 
	       
	       #headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
			 #'Accept-Encoding': 'gzip, deflate, br',
			 #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			 #'Content-Length': '10',
			 #'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
			 #'Cookie': 'session=aqq75kqhsmbegk00crvocv86t2',
			 #'Host': 'raui.ru',
			 #'Referer': task.url,
			 #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0', 
			 #'X-Requested-With': 'XMLHttpRequest'}
	       #g2 = grab.clone(headers=headers,proxy_auto_change=True)
	  
	      
	       #try:               
		    ##time.sleep(1)
		    #g2.request(post=[('id', id_phone)],headers=headers,url=phone_url)
		    
		    ##print g2.response.body
		    ##phone =  re.sub('[^\d\+]','',re.findall('em class=(.*?)/em>',g2.response.body)[0]) 
		    #phone =  re.sub('[^\d]','',g2.doc.json["contacts"]["phone"])
		    #print 'Phone-OK'
		    #del g2
	       #except (IndexError,TypeError,GrabConnectionError,GrabNetworkError,GrabTimeoutError):
		    #del g2
		    #phone = random.choice(list(open('../phone.txt').read().splitlines()))
		    
     
			 
	       
							
		    
	       projects = {'url': task.url,
		           'rayon': ray,
		           'sub': sub,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
		           'trassa': trassa,
		           'udal': udal,
		           'segment': seg,
		           'cena': price,
		           'plosh':plosh,
		           'linii': lini,
		           'internet':inet,
		           'uslugi':usl,
		           'parkov': park,
		           'cena_za': cena_za,
		           'ohrana':ohrana,
		           'phone':phone[:11],
		           'gaz': gaz,
		           'voda': voda,
		           'kanaliz': kanal,
		           'electr': elek,
		           'vhodd':vxod,
		           'lifft':lift,
		           'teplo': teplo.replace(udal+' ',''),
		           'potolok':pot,
		           'opis':opis,
		           'operazia':self.oper,
		           'sos': otd,
		           'shir': lat,
		           'dol': lng,	              
		           'data':data.strftime('%d.%m.%Y')}
	       
	       
	       
	       yield Task('write',project=projects,grab=grab)
		 
	  def task_write(self,grab,task):
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['trassa']
	       print  task.project['udal']
	       print  task.project['segment']
	       print  task.project['cena']
	       print  task.project['plosh']
	       print  task.project['cena_za']
	       print  task.project['shir']
	       print  task.project['dol']
	       print  task.project['parkov']
	       print  task.project['ohrana']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       #print  self.phone
	       print  task.project['data']
	       print  task.project['teplo']
	       
	       #global result
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 8, task.project['teritor'])
	       self.ws.write(self.result, 10, task.project['ulica'])
	       self.ws.write(self.result, 16, task.project['segment'])
	       self.ws.write(self.result, 15, task.project['trassa'])
	       self.ws.write(self.result, 9, task.project['udal'])
	       self.ws.write(self.result, 13 , task.project['dom'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 23, task.project['ohrana'])
	       self.ws.write(self.result, 17, task.project['cena_za'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 19, u'RAUI')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 22, task.project['gaz'])
	       self.ws.write(self.result, 26, task.project['voda'])
	       self.ws.write(self.result, 27, task.project['kanaliz'])
	       self.ws.write(self.result, 24, task.project['teplo'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 33, task.project['electr'])
	       self.ws.write(self.result, 28, task.project['operazia'])
	       self.ws.write(self.result, 34, task.project['shir'])
	       self.ws.write(self.result, 35, task.project['dol'])
	       self.ws.write(self.result, 37, task.project['parkov'])
	       self.ws.write(self.result, 38, task.project['uslugi'])
	       self.ws.write(self.result, 40, task.project['internet'])
	       self.ws.write(self.result, 41, task.project['linii'])
	       self.ws.write(self.result, 43, task.project['sos'])
	       self.ws.write(self.result, 45, task.project['vhodd'])
	       self.ws.write(self.result, 47, task.project['lifft'])
	       self.ws.write(self.result, 48, task.project['potolok'])
	       print('*'*50)
	       #print task.sub
	       
	       print 'Ready - '+str(self.result)#+'/'+str(self.num)+'0'
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print  task.project['operazia']
	       print('*'*50)	       
	       self.result+= 1
		    
		    
		    
	       if self.result > 5000:
		    self.stop()
     
	  
     bot = Raui_Com(thread_number=5,network_try_limit=1000)
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
    









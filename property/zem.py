#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab import Grab
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError
import re
import time
import os
import math
from sub import conv
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



i = 0
l= open('Links/Zem_Prod.txt').read().splitlines()

page = l[i]
oper = u'Продажа'

#g.proxylist.load_file(path='/home/oleg/Proxy/tipa.txt',proxy_type='http')



while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Theproperty_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       for p in range(1,21):
		    try:
			 time.sleep(2)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 print g.doc.code
			 if g.doc.code ==200:
			      self.num = re.sub('[^\d]','',g.doc.select(u'//p[@class="quant"]').text().split(', ')[0])
			      self.pag = int(math.ceil(float(int(self.num))/float(15)))
			      self.dt = g.doc.select(u'//p[@class="current-city"]/a').text().replace(',','').replace(u' г','')
			      print self.dt,self.num
			      link_sub = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+self.dt
			      g.go(link_sub)
			      #self.sub = g.doc.json["response"]["GeoObjectCollection"]["featureMember"][0]["GeoObject"]["metaDataProperty"]["GeocoderMetaData"]["Address"]["Components"][2]["name"]
			      self.sub = g.doc.rex_text(u'AdministrativeAreaName":"(.*?)"')
			      print self.sub,self.pag
			      del g
			      break
		    except(GrabTimeoutError,GrabNetworkError,IndexError,KeyError,AttributeError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
	            self.sub = ''
	            self.pag = 0
		    #del g
	       if self.pag == 0:
		    self.stop()

	       self.workbook = xlsxwriter.Workbook(u'zem/Theproperty_%s' % self.sub + u'_Земля_'+oper+str(i+1) + '.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Mlsn_Земля')
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
	       self.ws.write(0, 30, u"ТИП_СДЕЛКИ")
	       self.ws.write(0, 31, u"АДРЕС_САЙТА_ПРОДАВЦА")
	       self.ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
	       self.ws.write(0, 33, u"ШИРОТА_ИСХ")
	       self.ws.write(0, 34, u"ДОЛГОТА_ИСХ")

	       self.result= 1





	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'?page=%d'%x,refresh_cache=True,network_try_count=100)



	  def task_post(self,grab,task):
	       links = grab.doc.select(u'//p[@class="address"]/a')
	       for elem in links:
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)

	  def task_item(self, grab, task):
	       try:
		    ray = grab.doc.select(u'//dt[contains(text(),"Район")]/following-sibling::dd[1]').text()
	       except IndexError:
		    ray = ''
	       try:
		    a = grab.doc.select(u'//h1').text().split(' — ')[1]
		    count = len(a.split(','))-1
		    if count == 3:
			 punkt= a.split(', ')[0]
		    elif count == 2:
			 punkt = a.split(', ')[2]
		    elif count == 1:
			 punkt = a.split(', ')[1]
		    else:
			 punkt=''
	       except IndexError:
		    punkt = ''
	       try:
		    ter= grab.doc.select(u'//dt[contains(text(),"Метро, ориентир")]/following-sibling::dd[1]').text()
	       except IndexError:
		    ter =''
	       try:
		    a1 = grab.doc.select(u'//h1').text().split(' — ')[1]
		    count1 = len(a1.split(','))-1
		    if count1 == 3:
			 uliza= a1.split(', ')[1]
		    elif count1 == 2:
			 uliza = a1.split(', ')[0]
		    elif count1 == 1:
			 uliza = a1.split(', ')[0]
		    else:
			 uliza=''
	       except IndexError:
	            uliza = ''
	       try:
		    a2 = grab.doc.select(u'//h1').text().split(' — ')[1]
		    count2 = len(a2.split(','))-1
		    if count2 == 3:
			 dom= re.sub('[^\d]','',a2.split(', ')[2])[:2]
		    elif count2 == 2:
			 dom = re.sub('[^\d]','',a2.split(', ')[1])[:2]
		    elif count2 == 1:
			 dom = ''#re.sub('[^\d]','',a2.split(', ')[0])[:2]
		    else:
			 dom=''
	       except IndexError:
	            dom = ''



	       try:
		    metro = grab.doc.select(u'//p[@id="priceMulti_3_0"]/strong[1]').text()
		 #print rayon
	       except IndexError:
		    metro = ''

	       try:
		    metro_min = grab.doc.select(u'//h1').text().split(u' — ')[1]
		 #print rayon
	       except IndexError:
		    metro_min = ''

	       try:
		    metro_tr = grab.doc.select(u'//span[@class="object_item_metro_comment"]').text().split(u'мин. ')[1]
	       except IndexError:
		    metro_tr = ''

	       try:
		    try:
		         price = grab.doc.select(u'//p[@id="priceMulti_-1_1"]/strong[1]').text()
		    except IndexError:
			 price = grab.doc.select(u'//div[@class="cleaner"]/preceding-sibling::p[@class="descr"][1]/strong[1]').text()
               except IndexError:
                    price = ''

	       try:
		    plosh_ob = grab.doc.select(u'//dt[contains(text(),"Площадь")]/following-sibling::dd[1]').text()
		  #print rayon
	       except IndexError:
		    plosh_ob = ''



	       try:
		    et = grab.doc.select(u'//th[contains(text(),"Газоснабжение")]/following-sibling::td').text()
		 #print price + u' руб'
	       except IndexError:
		    et = ''

	       try:
		    etagn = grab.doc.select(u'//dt[contains(text(),"Тип сделки")]/following-sibling::dd[1]').text()
		 #print price + u' руб'
	       except IndexError:
		    etagn = ''




	       try:
		    try:
			 opis = grab.doc.select(u'//h2[contains(text(),"Дополнительная информация")]/following-sibling::p[2]').text()
		    except IndexError:
			 opis = grab.doc.select(u'//h2[contains(text(),"Дополнительная информация")]/following-sibling::p').text()
	       except IndexError:
	            opis = ''

	       try:
		    phone = re.sub('[^\d\+\,]','',grab.doc.select(u'//p[@class="phone"]').text())
	       except IndexError:
		    phone = ''

	       try:
		    lico = grab.doc.select(u'//p[@class="name"]/a').text()
	       except IndexError:
		    lico = ''

	       try:
		    comp = grab.doc.select(u'//p[@class="company"]/a').text()
		 #print rayon
	       except IndexError:
		    comp = ''

	       try:
		    lat = grab.doc.select(u'//div[@id="objMap"]').attr('data-coords').split(', ')[0].replace('[','')
	       except IndexError:
	            lat = ''
	       try:
	            lng = grab.doc.select(u'//div[@id="objMap"]').attr('data-coords').split(', ')[1]
	       except IndexError:
	            lng = ''

	       try:
		    data = grab.doc.select(u'//p[@class="email"]/noindex/a[contains(@rel,"nofollow")]').attr('href')
	       except IndexError:
		    data = ''


	       projects = {'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
		           'metro': metro,
	                   'naz': metro_min,
		           'tran': metro_tr,
		           'cena': price,
		           'plosh_ob':plosh_ob,
		           'etach': et,
	                   'shir':lat,
	                   'dol':lng,
		           'etashost': etagn,
		           'opis':opis,
		           'url':task.url,
		           'phone':phone,
		           'lico':lico,
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
	       print  task.project['dom']
	       print  task.project['metro']
	       print  task.project['tran']
	       print  task.project['cena']
	       print  task.project['plosh_ob']
	       print  task.project['etach']
	       print  task.project['etashost']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['naz']

	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 3,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 6,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 11,task.project['metro'])
	       self.ws.write(self.result, 32,task.project['naz'])
	       #self.ws.write(self.result, 10,task.project['object'])
	       self.ws.write(self.result, 9,oper)
	       self.ws.write(self.result, 10, task.project['cena'])
	       #self.ws.write(self.result, 13, task.project['cena_m'])
	       #self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 12, task.project['plosh_ob'])
	       self.ws.write(self.result, 33, task.project['shir'])
	       self.ws.write(self.result, 34, task.project['dol'])
	       #self.ws.write(self.result, 18, task.project['plosh_com'])
	       self.ws.write(self.result, 15, task.project['etach'])
	       self.ws.write(self.result, 30, task.project['etashost'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'TheProperty.ru')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write_string(self.result, 31, task.project['data'])
	       self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))


	       print('*'*50)
	       print self.sub

	       print 'Ready - '+str(self.result)+'/'+self.num
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*50)
	       self.result+= 1




	       #if self.result > 10:
		    #self.stop()



     bot = Theproperty_Zem(thread_number=5,network_try_limit=1000)
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
     del bot

     i=i+1
     try:
          page = l[i]
     except IndexError:
          if oper == u'Продажа':
               i = 0
               l= open('Links/Zem_Arenda.txt').read().splitlines()
               page = l[i]
               oper = u'Аренда'
	  else:
               break




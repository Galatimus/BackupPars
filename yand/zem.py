#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
import os
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError
import re
import time
import random
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logging.basicConfig(level=logging.DEBUG)



i = 0
l= open('Links/Zemm.txt').read().splitlines()
page = l[i]





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Mag_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.workbook = xlsxwriter.Workbook(u'zem/Yandex_Земля_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet()
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
	       self.ws.write(0, 30, u"МЕСТОПОЛОЖЕНИЕ")
	       self.ws.write(0, 31, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 32, u"ШИРОТА_ИСХ")
	       self.ws.write(0, 33, u"ДОЛГОТА_ИСХ")
	       self.result= 1





	  def task_generator(self):
	       for x in range(25):
		    yield Task ('post',url=self.f+'kupit/uchastok/?page=%d'%x,refresh_cache=True,network_try_count=100)	       
	       #yield Task ('post',url=self.f+'kupit/uchastok/',refresh_cache=True,network_try_count=100)



	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//div[@class="OffersSerpItem__generalInfo"]/a'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       #yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)


	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//a[@class="Pager__radio-link"][contains(text(),"Следующая")]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print('*'*100)
		    print '!!!','NO PAGE NEXT','!!!'
		    print('*'*100)


	  def task_item(self, grab, task):
	       try:
		    ray = grab.doc.select(u'//div[@class="OfferHeader"]/ol/li[1]/a/span').text()
	       except IndexError:
		    ray = ''
	       try:
		    punkt= grab.doc.select(u'//div[@class="OfferHeader__address"]').text().split(', ')[0].replace(ray,'')
	       except IndexError:
		    punkt = ''
	       try:
		    try:
                         ter =  grab.doc.select(u'//div[@class="OfferHeader"]/ol/li/a/span[contains(text(),"район")]').text()
		    except IndexError:
			 ter =  grab.doc.select(u'//div[@class="OfferHeader"]/ol/li/a/span[contains(text(),"округ")]').text()
	       except IndexError:
		    ter =''
	       try:
		    uliza = grab.doc.select(u'//div[@class="OfferHeader__address"]').text()
	       except IndexError:
		    uliza = ''
	       try:
		    dom = grab.doc.select(u'//div[@class="offer-card-sub-location"]').text().split(', ')[1]
	       except IndexError:
		    dom = ''

	       try:
		    orentir = grab.doc.select(u'//div[@class="offer-card-sub-location"]').text().split(', ')[0]
	       except IndexError:
		    orentir = ''

	       try:
		    metro = grab.doc.select(u'//div[@class="OfferBaseInfo__text-info"][contains(text(),"сот")]').text()
	       except IndexError:
		    metro = ''

	       try:
		    metro_min = grab.doc.select(u'//ul[@class="ColumnsList OfferTechDescription__list"]/li[1]/div/p').text().split(' — ')[0]
	       except IndexError:
		    metro_min = ''

	       try:
		    metro_tr = grab.doc.rex_text(u'updateDate":"(.*?)T').replace('-','.')
	       except IndexError:
		    metro_tr = ''

	       try:
		    price = grab.doc.select(u'//span[@class="price"]').text()
	       except IndexError:
		    price = ''

	       try:
		    plosh_ob = grab.doc.select(u'//ul[@class="ColumnsList OfferTechDescription__list"]/li[2]/div/p').text().split(' — ')[0]
	       except IndexError:
		    plosh_ob = ''

	       try:
		    et = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Назначение земли")]/following-sibling::div').text()
		 #print price + u' руб'
	       except IndexError:
		    et = ''

	       try:
		    etagn = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Газ")]/following-sibling::div').text()
		 #print price + u' руб'
	       except IndexError:
		    etagn = ''
               try:
		    opis = grab.doc.select(u'//p[@class="OfferTextDescription__text"]').text()
	       except IndexError:
		    opis = ''

	       try:
		    phone = re.sub('[^\d\+]','',grab.doc.rex_text(u'phoneNumbers(.*?)redirectId'))
	       except IndexError:
		    phone = random.choice(list(open('../phone.txt').read().splitlines()))

	       try:
		    lico = grab.doc.select(u'//div[@class="AuthorBadge__name"]').text()
	       except IndexError:
		    lico = ''

	       try:
		    comp = grab.doc.select(u'//div[@class="AuthorBadge__category"]').text()
	       except IndexError:
		    comp = ''

	       try:
	            data = grab.doc.rex_text(u'creationDate":"(.*?)T').replace('-','.')
	       except IndexError:
	            data=''

	       try:
	            lat = grab.doc.rex_text(u'latitude":(.*?),')
	       except IndexError:
	            lat =''

	       try:
	            lng = grab.doc.rex_text(u'longitude":(.*?),')
	       except IndexError:
	            lng =''


	       projects = {'rayon': ray,
		           'punkt': punkt.replace(ter,''),
		           'teritor':orentir,
		           'ulica': uliza,
	                   'mesto': ter,
		           'dom': dom,
		           'metro': metro,
	                   'naz': metro_min,
		           'tran': metro_tr,
		           'cena': price,
		           'plosh_ob':plosh_ob,
		           'etach': et,
		           'etashost': etagn,
		           'opis':opis,
		           'url':task.url,
		           'phone':phone[:12],
	                   'shir': lat,
	                   'dol': lng,
		           'lico':lico,
		           'company':comp,
		           'data':data}



	       yield Task('write',project=projects,grab=grab)






	  def task_write(self,grab,task):
               if task.project['rayon'] <> '':
		    print('*'*50)
		    print  task.project['rayon']
		    print  task.project['punkt']
		    print  task.project['mesto']
		    print  task.project['teritor']
		    print  task.project['ulica']
		    print  task.project['dom']
		    print  task.project['metro']
		    print  task.project['naz']
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
		    print  task.project['tran'] 
		    print  task.project['shir']
		    print  task.project['dol']
     
     
		    self.ws.write(self.result, 0,task.project['rayon'])
		    self.ws.write(self.result, 2,task.project['punkt'])
		    self.ws.write(self.result, 6,task.project['teritor'])
		    self.ws.write(self.result, 30,task.project['ulica'])
		    self.ws.write(self.result, 8,task.project['dom'])
		    self.ws.write(self.result, 11,task.project['metro'])
		    self.ws.write(self.result, 12,task.project['naz'])
		    self.ws.write(self.result, 31,task.project['tran'])
		    self.ws.write(self.result, 9,u'Продажа')
		    self.ws.write(self.result, 10, task.project['cena'])
		    self.ws.write(self.result, 1, task.project['mesto'])
		    self.ws.write(self.result, 32, task.project['shir'])
		    self.ws.write(self.result, 33, task.project['dol'])
		    self.ws.write(self.result, 13, task.project['plosh_ob'])
		    #self.ws.write(self.result, 14, task.project['etach'])
		    #self.ws.write(self.result, 15, task.project['etashost'])
		    self.ws.write(self.result, 22, task.project['opis'])
		    self.ws.write(self.result, 23, u'Яндекс Недвижимость')
		    self.ws.write_string(self.result, 24, task.project['url'])
		    self.ws.write(self.result, 25, task.project['phone'])
		    self.ws.write(self.result, 26, task.project['lico'])
		    self.ws.write(self.result, 27, task.project['company'])
		    self.ws.write(self.result, 28, task.project['data'])
		    self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
     
     
		    print('*'*50)
		    print 'Ready - '+str(self.result)
		    print 'Tasks - %s' % self.task_queue.size()
		    print '***',i+1,'/',len(l),'***'
		    print('*'*50)
		    self.result+= 1
     
     
     
     
		    #if self.result > 10:
			 #self.stop()


     bot = Mag_Zem(thread_number=5,network_try_limit=1000)
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
     bot.workbook.close()
     time.sleep(2)
     print('Done')
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break

time.sleep(5)
os.system("/home/oleg/pars/yand/comm.py")


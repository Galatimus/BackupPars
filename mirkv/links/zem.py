#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError
import logging
from grab import Grab
import re
import math
import random
import os
import time
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






i = 0
l= open('links/Zem.txt').read().decode('cp1251').splitlines()
dc = len(l)
page = l[i]
oper = u'Продажа'

while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class MK_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       for p in range(1,21):
		    try:
			 time.sleep(3)
			 g = Grab(timeout=10, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 self.sub = g.doc.select(u'//div[@class="c-dropdown-button"]/a').text()
			 self.num = re.sub('[^\d]','',g.doc.select(u'//span[@class="count"]').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 print self.sub,self.num,self.pag
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError,ValueError):
			 print g.config['proxy'],'Change proxy'
			 #g.change_proxy()
			 continue
	       else:
		    self.sub = ''
		    self.pag = 0
		    self.num=0

	       self.workbook = xlsxwriter.Workbook(u'zem/Mirkvartir_%s' % bot.sub + u'_Земля_'+oper+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Mirkvartir_Земля')
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
	       self.ws.write(0, 31, u"ШИРОТА_ИСХ")
	       self.ws.write(0, 32, u"ДОЛГОТА_ИСХ")	       


	       self.result= 1



	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'?p=%d'%x,refresh_cache=True,network_try_count=100)


	  def task_post(self,grab,task):
	       #if grab.doc.select(u'//h2[@class="nearby-header"]').exists()==True:
                    #links = grab.doc.select(u'//h2[@class="nearby-header"]/preceding::div[@class="item"]/a[1]')
               #else:
	            #links = grab.doc.select(u'//div[@class="item"]/a[1]')

               for elem in grab.doc.select(u'//a[@class="offer-title"]'):
	            ur = grab.make_url_absolute(elem.attr('href'))
	            #print ur
	            yield Task('item', url=ur,refresh_cache=True,network_try_count=100)



	  def task_item(self, grab, task):

	       try:
		    ray = grab.doc.select(u'//p[@class="address"]/a[contains(text(),"р-н")]').text()
		  #print ray
	       except IndexError:
		    ray = ''
	       try:
		    punkt= grab.doc.select(u'//meta[@name="keywords"]').attr('content').split(', ')[3]
	       except IndexError:
		    punkt = ''

	       try:
		    ter= grab.doc.select(u'//div[@class="b-breadcrumbs"]/ul/li/a/span[contains(text(),"поселение")]').text()
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
		    trassa = grab.doc.select(u'//a[@class="m-highway"]').text()
		     #print rayon
	       except IndexError:
		    trassa = ''

	       try:
		    udal = grab.doc.select(u'//div[@class="place"]').text()
	       except IndexError:
		    udal = ''

	       try:
		    price = grab.doc.select(u'//div[@class="price m-all"]').text()
	       except IndexError:
		    price = ''

               try:
		    plosh = grab.doc.select(u'//span[contains(text(),"Площадь")]/following-sibling::strong').text()
	       except IndexError:
		    plosh = ''

               try:
		    vid = grab.doc.select(u'//span[contains(text(),"Тип участка")]/following-sibling::strong').text()
	       except IndexError:
		    vid = ''

	       try:
		    ohrana =  grab.doc.select(u'//span[contains(text(),"Безопасность")]/following-sibling::strong').text().replace(u'охрана',u'есть')
	       except IndexError:
		    ohrana =''
	       try:
		    z =  grab.doc.select(u'//span[contains(text(),"Коммуникации")]/following-sibling::strong').text()
		    if z.find(u'газ')>=0:
			 gaz='есть'
		    else:
			 gaz=''
	       except IndexError:
		    gaz =''
	       try:
		    v =  grab.doc.select(u'//span[contains(text(),"Коммуникации")]/following-sibling::strong').text()
		    if v.find(u'вода')>=0:
			 voda='есть'
		    else:
			 voda=''
	       except IndexError:
		    voda =''
	       try:
		    k =  grab.doc.select(u'//span[contains(text(),"Коммуникации")]/following-sibling::strong').text()
		    if k.find(u'канализация')>=0:
			 kanal='есть'
		    else:
			 kanal=''
	       except IndexError:
		    kanal =''
	       try:
		    lk =  grab.doc.select(u'//span[contains(text(),"Коммуникации")]/following-sibling::strong').text()
		    if lk.find(u'электричество')>=0:
			 elek='есть'
		    else:
			 elek=''
	       except IndexError:
		    elek =''
	       try:
		    teplo = grab.doc.select(u'//p[@class="address"]').text()
	       except IndexError:
		    teplo =''

	       try:
		    opis = grab.doc.select(u'//div[@class="l-object-description"]/p').text()
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
		    try:
		         data = grab.doc.select(u'//div[@class="l-object-dates"]/p[2]').text().split(u' размещено ')[1].split(u' в ')[0].replace(u'сегодня', (datetime.today().strftime('%d.%m.%Y')))
		    except IndexError:
			 data = grab.doc.select(u'//div[@class="dates"]').text().split(u' размещено ')[1].split(u' в ')[0].replace(u'сегодня', (datetime.today().strftime('%d.%m.%Y')))
	       except IndexError:
		    data = ''

	       try:
                    phone =  re.sub('[^\d\+]','',grab.doc.select(u'//a[@class="phone"]').text())
               except IndexError:
	            phone = ''

	       try:
		    lat = grab.doc.rex_text(u'"lat":(.*?),')
	       except IndexError:
	            lat =''
		    
	       try:
	            lng = grab.doc.rex_text(u'"lon":(.*?)}')
	       except IndexError:
	            lng =''
		    
	       try:
	            naz = grab.doc.select(u'//div[@class="price m-m2"]').text()
	       except IndexError:
	            naz = ''    



	       projects = {'url': task.url,
		           'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt.replace(ter,''),
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
		           'trassa': trassa,
		           'udal': udal.replace(trassa,''),
		           'cena': price,
		           'plosh':plosh,
		           'vid': vid,
		           'ohrana':ohrana,
		           'gaz': gaz,
		           'voda': voda,
	                   'phone':phone,
		           'kanaliz': kanal,
		           'electr': elek,
		           'teplo': teplo,
		           'opis':opis,
	                   'sotka': naz,
		           'lico':lico.replace(comp,''),
	                   'dol': lat,
	                   'shir': lng,	                    
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
	       print  task.project['trassa']
	       print  task.project['udal']
	       print  task.project['cena']
	       print  task.project['plosh']
	       print  task.project['ohrana']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['teplo']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['vid']
	       print  task.project['data']
	       print  task.project['dol']
	       print  task.project['shir']
	       print  task.project['sotka']


	       #global result
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 7, task.project['trassa'])
	       self.ws.write(self.result, 8, task.project['udal'])
	       self.ws.write(self.result, 9, oper)
	       self.ws.write(self.result, 10, task.project['cena'])
	       self.ws.write(self.result, 11, task.project['sotka'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 14, task.project['vid'])
	       self.ws.write(self.result, 15, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 18, task.project['electr'])
	       self.ws.write(self.result, 30, task.project['teplo'])
	       self.ws.write(self.result, 20, task.project['ohrana'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'MIRKVARTIR.RU')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 28, task.project['data'])
	       self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 32, task.project['shir'])
	       self.ws.write(self.result, 31, task.project['dol'])	       

	       print('*'*50)
	       print self.sub
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',dc,'***'
	       print oper
	       print('*'*50)

	       self.result+= 1



	       if self.result > 7000:
		    self.stop()


     bot = MK_Zem(thread_number=10,network_try_limit=1000)
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








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
from datetime import datetime,timedelta
import random
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



i = 0
l= open('links/zem_prod.txt').read().splitlines()

page = l[i]
oper = u'Продажа'





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Nedvizhka_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       #self.link =l[i]
	       for p in range(1,16):
		    try:
                         time.sleep(1)
			 g = Grab(timeout=5, connect_timeout=10)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                         g.go(self.f) 
			 #self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="total"]/p').text())
			 #self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 self.sub = g.doc.select(u'//div[@class="breadcrumbs"]/span[2]/a').text().replace('/','-')
			 print self.sub
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
		         print g.config['proxy'],'Change proxy'
		         g.change_proxy()
			 del g
		         continue
	       else:
	            self.sub = ''


	       self.workbook = xlsxwriter.Workbook(u'zem/Move_Земля_'+oper+str(i)+'.xlsx')
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
	       self.ws.write(0, 21, u"ПЛОЩАДЬ_В_КМ")
	       self.ws.write(0, 22, u"ОПИСАНИЕ")
	       self.ws.write(0, 23, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 24, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 25, u"ТЕЛЕФОН")
	       self.ws.write(0, 26, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 27, u"КОМПАНИЯ")
	       self.ws.write(0, 28, u"ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 30, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
	       self.ws.write(0, 32, u"ШИРОТА_ИСХ")
	       self.ws.write(0, 33, u"ДОЛГОТА_ИСХ")
	       self.result= 1






	  def task_generator(self):
	       #for x in range(1,self.pag+1):
                    #yield Task ('post',url=self.f+'?page=%d'%x+'&limit=20',refresh_cache=True,network_try_count=100)
               yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)


	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@class="search-item__title-link search-item__item-link"]'):
	            ur = grab.make_url_absolute(elem.attr('href'))
	            #print ur
	            yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	       
	       
	       
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//a[@class="pagination-block__paginate-next pagination-block__ctrl-title"]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print 'no_page'

	  def task_item(self, grab, task):
	       try:
		    ray = grab.doc.select(u'//div[@class="breadcrumbs"]/span/a[contains(@title, "район")]').text()
	       except IndexError:
		    ray = ''
	       try:
		    try:
			 punkt = grab.doc.select(u'//div[@class="breadcrumbs"]/span/a[contains(@title, "в г")]').text()
		    except IndexError:
			 punkt = grab.doc.select(u'//div[@class="breadcrumbs"]/span/a[contains(@title, "поселок")]').text()
	       except IndexError:
			 punkt = ''

	       ter =''


	       try:
		    try:
			 try:
			      uliza = grab.doc.select(u'//div[@class="breadcrumbs"]/span/a[contains(@title, "улице")]').text()
			 except IndexError:
			      uliza = grab.doc.select(u'//span[@class="geo-block__geo-info_no-link"][contains(text(),"улица")]').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//span[@class="geo-block__geo-info_no-link"][contains(text(),"проезде")]').text()
	       except IndexError:
	            uliza =''



	       try:
		    dom = re.sub('[^\d]','',grab.doc.select(u'//h1/span[@class="object-title_page-title_tail"]').text())
                    #dom = re.compile(r'[0-9]+$',re.S).search(dm).group(0)
	       except IndexError:
		    dom = ''

	       try:
		    orentir = grab.doc.select(u'//div[@class="object-place__address"]/following-sibling::div[@class="object-info"]/div').text().replace(u' на карте ',', ')
	       except IndexError:
		    orentir = ''

	       try:
		    metro = grab.doc.select(u'//div[contains(text(),"цена за сотку:")]/following-sibling::div').text()#.split('/')[1]
		 #print rayon
	       except IndexError:
		    metro = ''

	       try:
		    metro_min = grab.doc.select(u'//div[contains(text(),"Тип объекта:")]/following-sibling::div').text()
		 #print rayon
	       except IndexError:
		    metro_min = ''

	       try:
		    metro_tr = grab.doc.select(u'//div[@class="object-place__address"]').text()
	       except IndexError:
		    metro_tr = ''


               try:
	            lat = grab.doc.select(u'//script[@type="text/javascript"][contains(text(),"coordsCenterTile")]').text().split('coordsCenterTile=')[1].split('];')[0].replace('[','').split(',')[0]
	       except IndexError:
	            lat =''

	       try:
	            lng = grab.doc.select(u'//script[@type="text/javascript"][contains(text(),"coordsCenterTile")]').text().split('coordsCenterTile=')[1].split('];')[0].replace('[','').split(',')[1]
	       except IndexError:
	            lng =''


	       try:
		    price = grab.doc.select(u'//div[contains(text(),"Цена:")]/following-sibling::div').text()
		 #print price + u' руб'
	       except IndexError:
		    price = ''

	       try:
		    plosh_ob = grab.doc.select(u'//div[contains(text(),"Площадь участка:")]/following-sibling::div').text()
	       except IndexError:
		    plosh_ob = ''

	       try:
		    plosh = grab.doc.select(u'//div[contains(text(),"Общая площадь:")]/following-sibling::div').text()
	       except IndexError:
	            plosh = ''

	       try:
		    et = grab.doc.select(u'//div[@class="object-place"]').text().split(u' на карте')[0].replace(u'Расположение ','')
		 #print price + u' руб'
	       except IndexError:
		    et = ''

	       try:
		    opis = grab.doc.select(u'//h4[contains(text(),"Описание")]/following-sibling::div[1]').text().replace('=','')
	       except IndexError:
		    opis = ''


	       try:
		    lico = grab.doc.select(u'//div[@class="block-user__name"]').text()
	       except IndexError:
		    lico = ''

	       try:
		    comp = grab.doc.select(u'//div[@class="block-user__agency"]').text().replace(u'Риелтор','')
		 #print rayon
	       except IndexError:
		    comp = ''



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
	                   'plosh':plosh,
		           'etach': et,
		           'opis':opis,
	                   'oren':orentir.replace(u' на карте',''),
	                   'shir':lat,
	                   'dol':lng,
		           'url':task.url,
		           'lico':lico,
		           'company':comp}

	       try:
		    link = task.url+'print/'
		    yield Task('phone',url=link,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
	            yield Task('phone',grab=grab,project=projects)


	  def task_phone(self, grab, task):
	       try:
		    phone= re.sub('[^\d\,]','',grab.doc.select(u'//div[@class="phone"]').text())
	       except IndexError:
		    phone = ''
	       try:
		    data1=  grab.doc.select(u'//div[@class="tech-info"]/div[2]/span').text().split(' ')[1]
	       except IndexError:
		    data1 =''
	       try:
		    data = grab.doc.select(u'//div[@class="tech-info"]/div[1]/span').text().split(' ')[1]
	       except IndexError:
		    data = ''

	       project2 ={'phone':phone,
	                  'dataraz': data,
	                  'dataob':data1}


	       yield Task('write',project=task.project,proj=project2,grab=grab)






	  def task_write(self,grab,task):
	       #if task.project['phone']<>'':
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['metro']
	       print  task.project['naz']
	       print  task.project['cena']
	       print  task.project['plosh_ob']
	       print  task.project['etach']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.proj['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.proj['dataraz']
               print  task.proj['dataob']
	       print  task.project['tran']
	       print  task.project['oren']
	       print  task.project['plosh']


	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 11,task.project['metro'])
	       self.ws.write(self.result, 13,task.project['naz'])
	       self.ws.write(self.result, 31,task.project['tran'])
	       self.ws.write(self.result, 9,oper)
	       self.ws.write(self.result, 10, task.project['cena'])
	       self.ws.write(self.result, 32, task.project['shir'])
	       self.ws.write(self.result, 33, task.project['dol'])
	       self.ws.write(self.result, 12, task.project['plosh_ob'])
	       self.ws.write(self.result, 6, task.project['oren'])
	       self.ws.write(self.result, 21, task.project['plosh'])
	       #self.ws.write(self.result, 18, task.project['plosh_com'])
	       #self.ws.write(self.result, 31, task.project['etach'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'MOVE.RU')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.proj['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 28, task.proj['dataraz'])
               self.ws.write(self.result, 29, task.proj['dataob'])
	       self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))


	       print('*'*50)
	       print self.sub

	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*50)
	       self.result+= 1

	       #if str(self.result) == str(self.num):
		    #self.stop()


	       if self.result > 15000:
		    self.stop()

     bot = Nedvizhka_Zem(thread_number=7,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     #command = 'mount -a'
     #os.system('echo %s|sudo -S %s' % ('1122', command))
     #time.sleep(2)
     bot.workbook.close()
     print('Done')
     i=i+1
     try:
          page = l[i]
     except IndexError:
	  if oper == u'Продажа':
	       i = 0
	       l= open('links/zem_arenda.txt').read().splitlines()
	       dc = len(l)
	       page = l[i]
	       oper = u'Аренда'
	  else:
               break




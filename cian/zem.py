#! /usr/bin/env python
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

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'0001-0062_00_У_005-0002_CIAN.xlsx')




class Cian_Zem(Spider):
     def prepare(self):

	  self.ws = workbook.add_worksheet('Cian')
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
	  l= open('cian_zem.txt').read().splitlines()
	  self.dc = len(l)
	  print self.dc
	  for line in l:
	       yield Task ('item',url=line,refresh_cache=True,network_try_count=100)

     def task_item(self, grab, task):

	  try:
	       sub = grab.doc.select(u'//address[contains(@class,"address")]').text().split(', ')[0]
	  except IndexError:
	       sub = ''

	  try:
	       ray = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"р-н ")]').text()
	     #print ray
	  except DataNotFound:
	       ray = ''
	  try:
	       if  grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"р-н ")]').exists()==True:
		    punkt= grab.doc.select(u'//address[contains(@class,"address")]').text().split(', ')[2]
	       else:
		    punkt= grab.doc.select(u'//address[contains(@class,"address")]').text().split(', ')[1]
	  except IndexError:
	       punkt = ''

	  try:
	       #if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text()," район")]').exists()==True:
	       ter= grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text(),"район ")]').text()#.split(', ')[3].replace(u'улица','')
	       #else:
		    #ter= ''#grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[2]
	  except IndexError:
	       ter =''

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
				   uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"б-р")]').text()
			 except IndexError:
			      uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"бул.")]').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"проезд")]').text()
	       except IndexError:
		    uliza = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(text(),"наб.")]').text()
	  except IndexError:
	       uliza = ''
	  try:
	       dom = grab.doc.select(u'//address[contains(@class,"address")]/a[contains(@href,"house")]').text()
	  except DataNotFound:
	       dom = ''

	  try:
	       trassa = grab.doc.select(u'//a[contains(text(),"шоссе")]').text()
		#print rayon
	  except DataNotFound:
	       trassa = ''

	  try:
	       udal = grab.doc.select(u'//span[@class="highway_distance--1Gy1i"]').text()
	  except DataNotFound:
	       udal = ''

	  try:
	       price = grab.doc.select(u'//span[@itemprop="price"]').text()
	  except DataNotFound:
	       price = ''

	  try:
	       try:
		    plosh = grab.doc.select(u'//span[contains(text(),"Площадь участка")]/following-sibling::span[1]').text()#.replace(u'участок ','')
	       except IndexError:
		    plosh = grab.doc.select(u'//dt[contains(text(),"Площадь:")]/following-sibling::ddd').text()
	  except IndexError:
	       plosh = ''

	  try:
	       vid = grab.doc.select(u'//span[contains(text(),"Статус участка")]/following-sibling::span[1]').text()
	  except DataNotFound:
	       vid = ''


	  try:
	       ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
	  except DataNotFound:
	       ohrana =''
	  try:
	       gaz = grab.doc.select(u'//th[contains(text(),"Газ:")]/following-sibling::td').text()
	  except DataNotFound:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//th[contains(text(),"Водоснабжение:")]/following-sibling::td').text()
	  except DataNotFound:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//th[contains(text(),"Канализация:")]/following-sibling::td').text()
	  except DataNotFound:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//th[contains(text(),"Электричество:")]/following-sibling::td').text()
	  except DataNotFound:
	       elek =''
	  try:
	       teplo = grab.doc.select(u'//address[contains(@class,"address")]').text().replace(u'На карте','')
	  except IndexError:
	       teplo =''

	  try:
	       if 'sale' in task.url:
		    oper = u'Продажа'
	       elif 'rent' in task.url:
		    oper = u'Аренда'
	  except IndexError:
	       oper = ''


	  try:
	       opis = grab.doc.select(u'//p[@itemprop="description"]').text()#.split(u'Показать телефон')[0]
	  except IndexError:
	       opis = ''

	  try:
	       phone = grab.doc.rex_text(u'href="tel:(.*?)"')
	  except IndexError:
	       phone = ''
	  try:
	       lat = grab.doc.rex_text(u'center=(.*?)&').split('%2C')[0]
          except IndexError:
	       lat =''

          try:
	       lng = grab.doc.rex_text(u'center=(.*?)&').split('%2C')[1]
	  except IndexError:
	       lng =''

	  try:
	       try:
		    lico = grab.doc.select(u'//h3[@class="realtor-card__title"][contains(text(),"Представитель: ")]').text().replace(u'Представитель: ','')
	       except IndexError:
		    lico = grab.doc.select(u'//a[contains(@href,"agents")]/h2').text()
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
	       data = re.sub(u'[^\d\-]','',grab.doc.rex_text(u'editDate(.*?)T')).replace('-','.')
	    #print data
	  except DataNotFound:
	       data = ''




	  projects = {'url': task.url,
                      'sub': sub,
                      'rayon': ray,
                      'punkt': punkt,
                      'teritor': ter,
                      'ulica': uliza,
                      'dom': dom,
                      'trassa': trassa,
                      'udal': udal,
                      'cena': price,
                      'plosh':plosh,
                      'vid': vid,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'phone':phone.replace(u'79311111111',''),
                      'lico':lico,
	              'dol': lat,
	              'shir': lng,
                      'company':comp,
                      'data':data,
                      'oper':oper
                      }

	  yield Task('write',project=projects,grab=grab)

     def task_write(self,grab,task):
	  if task.project['sub'] <> '':
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
	       print  task.project['vid']
	       print  task.project['ohrana']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['teplo']
	       #print  task.project['oper']

	       #global result
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 7, task.project['trassa'])
	       self.ws.write(self.result, 8, task.project['udal'])
	       self.ws.write(self.result, 9, task.project['oper'])
	       self.ws.write(self.result, 10, task.project['cena'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 14, task.project['vid'])
	       self.ws.write(self.result, 15, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 18, task.project['electr'])
	       self.ws.write(self.result, 30, task.project['teplo'])
	       self.ws.write(self.result, 20, task.project['ohrana'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'ЦИАН')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 28, task.project['data'])
	       self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 32, task.project['shir'])
	       self.ws.write(self.result, 31, task.project['dol'])
	       print('*'*50)
	       #print task.sub

	       print 'Ready - '+str(self.result)+'/'+str(self.dc)
	       print 'Tasks - %s' % self.task_queue.size()
	       print  task.project['oper']
	       print('*'*50)
	       self.result+= 1

	       #if self.result > 20:
		    #self.stop()


bot = Cian_Zem(thread_number=5,network_try_limit=1000)
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
workbook.close()
print('Done')










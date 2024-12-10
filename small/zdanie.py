#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import os
import math
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






   
     
class move_Com(Spider):
     def prepare(self):
	  self.workbook = xlsxwriter.Workbook(u'comm/0001-0002_00_C_001-0227_ZDANIE.xlsx')
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
	  self.ws.write(0, 36, u"ЭЛЕКТРОСНАБЖЕНИЕ")
	  self.ws.write(0, 37, u"СИСТЕМА ВЕНТИЛЯЦИИ")
	  self.ws.write(0, 38, u"ОХРАНА")
	  self.ws.write(0, 39, u"КОНДИЦИОНИРОВАНИЕ")
	  self.ws.write(0, 40, u"ПЛАНИРОВКА")
	  self.ws.write(0, 41, u"ВЫСОТА ПОТОЛКОВ")
	  self.result= 1
	  
	   
	   
	   
	   
     def task_generator(self):
	  yield Task ('next',url='https://zdanie.info/аренда',refresh_cache=True,network_try_count=100)
	  yield Task ('next',url='https://zdanie.info/продажа',refresh_cache=True,network_try_count=100)
	  
	  
     def task_next(self,grab,task):
	  for el in grab.doc.select(u'//div[@class="checkbox"]/a'):
	       urr = grab.make_url_absolute(el.attr('href'))  
	       #print urr
	       yield Task('post', url=urr,refresh_cache=True,network_try_count=100)

     
   
     def task_post(self,grab,task):
	  for elem in grab.doc.select('//h2/a[contains(@href,"object")]'):
	       ur = grab.make_url_absolute(elem.attr('href'))
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	  yield Task('page', grab=grab,refresh_cache=True,network_try_count=100)   
       
     def task_page(self,grab,task):
	  try:
	       pg = grab.doc.select(u'//li[@class="pager-arr-r"]/a')
	       u = grab.make_url_absolute(pg.attr('href'))
	       yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	  except IndexError:
	       print 'no_page'
	       
     def task_item(self, grab, task):
	  

	  try:
	       mesto = grab.doc.select(u'//div[@class="object-standart-addr text-upper mb-5 fz10 cl-blue"]').text().replace(' › ',', ')
	  except IndexError:
	       mesto =''	 
	    
	  try:
	       tip = grab.doc.select(u'//h2[@class="section-title no-pd-l"]/text()').text()
	  except IndexError:
	       tip = ''
	  try:
	       naz = grab.doc.select(u'//h3/b[contains(text(),"Возможные варианты использования")]/following::ul[1]/li/a').text()
	  except IndexError:
	       naz =''
	  try:
	       klass =  grab.doc.select(u'//title').text().split(u'класса ')[1].split(u' с ')[0]
	  except IndexError:
	       klass = ''
	  try:
	       price = grab.doc.select(u'//td[@data-price-col="total"]/span[2]').text()+u' р.'    
	  except IndexError:
	       price =''
	  try: 
	       plosh = grab.doc.select(u'//tbody/tr/td[@class="cl-red"]').text()
	  except IndexError:
	       plosh=''
	  try:
	       ohrana = grab.doc.select(u'//li[contains(text(),"Этаж")]').text().split(' — ')[1].split(' из ')[0]
	  except IndexError:
	       ohrana =''
	  try:
	       gaz =  grab.doc.select(u'//li[contains(text(),"Этаж")]').text().split(' — ')[1].split(' из ')[1]
	  except IndexError:
	       gaz =''
	  try:
	       voda =  grab.doc.select(u'//li[contains(text(),"год постройки")]').number()
	  except IndexError:
	       voda =''
	  try:
	       kanal =  grab.doc.rex_text(u'create_time":"(.*?) ').replace('-','.')
	  except IndexError:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//div[@class="object-standart-addr text-upper mb-5 fz10 cl-blue"]/a[contains(text(), "м ")]').text()
	  except IndexError:
	       elek =''
	  try:
	       teplo = grab.doc.rex_text(u'update_time":"(.*?) ').replace('-','.')
	  except IndexError:
	       teplo =''
	  #time.sleep(1)
	  try:
	       opis = grab.doc.select(u'//div[contains(@class,"desription")]').text() 
	  except IndexError:
	       opis = ''
	  try:
	       lico = grab.doc.select(u'//i[contains(@class,"user fz17")]/following::td[1]/strong').text()
	  except IndexError:
	       lico = ''
	  try:
	       comp = grab.doc.select(u'//td[contains(text(),"Компания:")]/following-sibling::td[1]').text()
	  except IndexError:
	       comp = ''
	  
	  try: 
	       data = grab.doc.rex_text(u'geo_lat(.*?),')
	  except IndexError:
	       data=''
	       
	  try:
	       lng = grab.doc.rex_text(u'geo_lon(.*?),')
          except IndexError:
	       lng =''	       
	  try:
	       phone = re.sub('[^\d\;\+]','',grab.doc.select(u'//td[@class="border-b"]/strong[contains(text(),"+")]').text())
	  except IndexError:
	       phone = ''
	       
	  try:
	       if 'Продажа' in tip:
	            oper = u'Продажа' 
	       elif 'Аренда' in tip:
	            oper = u'Аренда'
	       else:
	            oper = grab.doc.select(u'//div[@class="breadcrumbs cl-grey"]/a[contains(text(),"регионах России")][1]').text().split(' ')[0]
          except IndexError:
	       oper = ''           
	  
	  try:
	       mesto1 = grab.doc.select(u'//li[contains(text(),"Охрана")]').text().split(' — ')[1]
	  except IndexError:
	       mesto1 =''    
	  try:
	       mesto2 = grab.doc.select(u'//li[contains(text(),"Система кондиционирования")]').text().split(' — ')[1]
	  except IndexError:
	       mesto2 =''
	  try:
	       mesto3 = grab.doc.select(u'//li[contains(text(),"Высота потолков")]').text().split(' — ')[1]
	  except IndexError:
	       mesto3 =''    
	  try:
	       mesto4 = grab.doc.select(u'//li[contains(text(),"Электроснабжение")]').text().split(' — ')[1]
	  except IndexError:
	       mesto4 =''
	       
	  try:
	       mesto5 = grab.doc.select(u'//li[contains(text(),"Система вентиляции")]').text().split(' — ')[1]
	  except IndexError:
	       mesto5 =''       

	  try:
	       mesto6 = grab.doc.select(u'//li[contains(text(),"планировка")]').text()
	  except IndexError:
	       mesto6 =''       
   
	  projects = {'adress': mesto,
                      'tip':tip.split(' ')[1].replace(',',''),
                      'naz':naz,
                      'klass': klass,
                      'cena': price,
                      'plosh': plosh,
	              'operacia': oper,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis': opis,
	              'dol': re.sub('[^\d\.]','',lng),
                      'url': task.url,
                      'phone': phone,
                      'lico':lico,
                      'company': comp.replace(u'Не указано',''),
	              'zag': tip,
	              'adress1': mesto1,
	              'adress2': mesto2,
	              'adress3': mesto3,
	              'adress4': mesto4,
	              'adress5': mesto5,
	              'adress6': mesto6,              
                      'data':re.sub('[^\d\.]','',data)}
	  
	  try:
	       link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+mesto
	       yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=100)
	  except IndexError:
	       yield Task('adres',grab=grab,project=projects)
	       
		      
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
	  #time.sleep(1)
	  print('*'*100)	       
	  print  task.proj['sub']
	  print  task.proj['punkt']  
	  print  task.proj['teritor']
	  print  task.proj['ulica']
	  print  task.proj['dom']
	  print  task.project['tip']
	  print  task.project['naz']
	  print  task.project['klass']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanaliz']
	  print  task.project['electr']
	  print  task.project['teplo']
	  print  task.project['ohrana']
	  print  task.project['opis']
	  print  task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company'] 
	  print  task.project['adress']
	  print  task.project['zag']
	  print  task.project['data']
	  print  task.project['dol']
	 
	 
     
	  
	  

	  self.ws.write(self.result, 0, task.proj['sub'])
	  self.ws.write(self.result, 24, task.project['adress'])
	  self.ws.write(self.result, 3, task.proj['teritor'])
	  self.ws.write(self.result, 2, task.proj['punkt'])
	  self.ws.write(self.result, 4, task.proj['ulica'])
	  self.ws.write(self.result, 5, task.proj['dom'])
	  self.ws.write(self.result, 7, task.project['tip'])
	  self.ws.write(self.result, 9, task.project['naz'])
	  self.ws.write(self.result, 10, task.project['klass'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  self.ws.write(self.result, 15, task.project['ohrana'])
	  self.ws.write(self.result, 16, task.project['gaz'])
	  self.ws.write(self.result, 17, task.project['voda'])
	  self.ws.write(self.result, 29, task.project['kanaliz'])
	  self.ws.write(self.result, 26, task.project['electr'])
	  self.ws.write(self.result, 33, task.project['zag'])
	  self.ws.write(self.result, 18, task.project['opis'])
	  self.ws.write(self.result, 19, u'Zdanie.Info')
	  self.ws.write_string(self.result, 20, task.project['url'])
	  self.ws.write(self.result, 21, task.project['phone'])
	  self.ws.write(self.result, 22, task.project['lico'])
	  self.ws.write(self.result, 23, task.project['company'])
	  self.ws.write(self.result, 34, task.project['data'])
	  self.ws.write(self.result, 30, task.project['teplo'])
	  self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 28, task.project['operacia'])
	  self.ws.write(self.result, 35, task.project['dol'])
	  self.ws.write(self.result, 38, task.project['adress1'])  
	  self.ws.write(self.result, 39, task.project['adress2'])
	  self.ws.write(self.result, 41, task.project['adress3'])
	  self.ws.write(self.result, 36, task.project['adress4'])
	  self.ws.write(self.result, 37, task.project['adress5'])
	  self.ws.write(self.result, 40, task.project['adress6'])  
	  
	  print('*'*10)
	  #print self.sub
	  print 'Ready - '+str(self.result)
	  logger.debug('Tasks - %s' % self.task_queue.size()) 
	  print task.project['operacia']
	  print('*'*10)
	  self.result+= 1


bot = move_Com(thread_number=5, network_try_limit=2000)
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
time.sleep(5)
os.system("/home/oleg/pars/small/tomsk_zem.py")
    
       
     
     
     
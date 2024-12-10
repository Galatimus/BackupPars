#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import os
import random
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
	  self.workbook = xlsxwriter.Workbook(u'comm/0001-0113_00_C_001-0245_ARNTOR.xlsx')
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
	  yield Task ('next',url='https://www.arendator.ru/towns/',network_try_count=100)
	  
	  
     def task_next(self,grab,task):
	  for el in grab.doc.select(u'//div[@class="regions-list__item col-4"]/a'):
	       urr = grab.make_url_absolute(el.attr('href'))#.split('.')[0].replace('http','https')+'.arendator.ru/'  
	       #print urr
	       yield Task('post', url=urr+'/offers/spaces/sale/',network_try_count=100)
	       yield Task('post', url=urr+'/offers/spaces/',network_try_count=100)
	  yield Task('post', url='https://www.arendator.ru/offers/spaces/sale/',network_try_count=100)
	  yield Task('post', url='https://www.arendator.ru/offers/spaces/',network_try_count=100)     
   
     def task_post(self,grab,task):
	  for elem in grab.doc.select('//a[contains(@class,"objects-list__box")]'):
	       ur = grab.make_url_absolute(elem.attr('href'))
	       #print ur
	       yield Task('item', url=ur,network_try_count=100)
	  yield Task('page', grab=grab,network_try_count=100)   
       
     def task_page(self,grab,task):
	  try:
	       pg = grab.doc.select(u'//a[contains(@title,"Следующая страница")]')
	       u = grab.make_url_absolute(pg.attr('href'))
	       yield Task ('post',url= u,network_try_count=100)
	  except IndexError:
	       print 'no_page'
	       
     def task_item(self, grab, task):
	  try:
	       sub = grab.doc.select(u'//i[@class="icon-point"]/following-sibling::text()').text()
	  except IndexError:
	       sub =''	  
	  try:
	       mesto = grab.doc.select(u'//div[@class="indexcard-info-params__adress"]').text()
	  except IndexError:
	       mesto =''
	       
	  try:
	       punkt = grab.doc.select(u'//div[@class="indexcard-info-params__adress"]').text().split(',')[0].split(' (')[0]
	  except IndexError:
	       punkt = ''	       
	   
	  try:
	       ter= grab.doc.select(u'//h1/a[contains(@href,"objects")]').text()
	  except IndexError:
	       ter =''
	  try:
	       uliza= grab.doc.rex_text(u'street_title(.*?)street_id')[3:][:-3].decode("unicode_escape") 
	  except IndexError:
	       uliza =''
	  try:
	       dom = grab.doc.select(u'//ul[@class="indexcard__breadcrumbs breadcrumbs"]/li[2]/a/span').text().split(' ')[0]
	  except (IndexError,AttributeError):
	       dom = ''
	    
	  try:
	       tip = grab.doc.select(u'//ul[@class="indexcard__breadcrumbs breadcrumbs"]/li[3]/a/span').text()
	  except IndexError:
	       tip = ''
	  try:
	       naz = grab.doc.select(u'//ul[@class="indexcard__breadcrumbs breadcrumbs"]/li[2]/a/span').text()
	  except IndexError:
	       naz =''
	  try:
	       klass =  grab.doc.select(u'//div[contains(text(),"Класс офисного здания")]/following-sibling::div').text()
	  except IndexError:
	       klass = ''
	  try:
	       
	       price = grab.doc.select(u'//div[@class="indexcard-info-params-box"][1]/div[1]').text()
	  except IndexError:
	       price =''
	  try: 
	       plosh = grab.doc.select(u'//div[@class="indexcard-info-params-box"][2]/div[1]').text().split(': ')[1]
	  except IndexError:
	       plosh=''
	  try:
	       ohrana = grab.doc.select(u'//div[contains(text(),"Этаж")]/following-sibling::div').text()
	  except IndexError:
	       ohrana =''
	  try:
	       gaz =  grab.doc.select(u'//div[@class="indexcard-info-params-metro"]/div').text().split(' - ')[0]
	  except IndexError:
	       gaz =''
	  try:
	       voda =  grab.doc.select(u'//div[@class="indexcard-info-params-metro"]/div').text().split(' - ')[1]
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.rex_text(u'datePublished(.*?)description')
	  except IndexError:
	       kanal =''
	  try:
	       elek = grab.doc.rex_text(u'lat(.*?),')
	  except IndexError:
	       elek =''
	  try:
	       teplo = grab.doc.rex_text(u'lon(.*?),')
	  except IndexError:
	       teplo =''

	  try:
	       opis = grab.doc.select(u'//div[@class="indexcard-desc__info"]').text() 
	  except IndexError:
	       opis = ''
	  try:
	       lico = grab.doc.rex_text(u'provider(.*?)datePublished').split('name')[1][3:][:-4]
	  except IndexError:
	       lico = ''
	  try:
	       comp = grab.doc.select(u'//title').text()
	  except IndexError:
	       comp = ''
	  
	  try: 
	       data = grab.doc.rex_text(u'Год постройки: (.*?)}')
	  except IndexError:
	       data=''
	       
	  
	  try:
	       phone = '+'+ grab.doc.rex_text(u'contact_val(.*?)contacts_types_ids').split('+')[1]
	  except IndexError:
	       phone = random.choice(list(open('../phone.txt').read().splitlines()))
	       

   
	  projects = {'sub': sub,
                     'adress': mesto,
                      'terit':ter, 
                      'punkt':punkt, 
                      'ulica':uliza,
                      'dom':dom,
                      'tip':tip.replace(dom,'')[1:],
                      'naz':naz.replace(dom,'')[1:],
                      'klass': klass,
                      'cena': price,
                      'plosh': plosh,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': re.sub('[^\d\-]','',kanal).replace('-','.'),
                      'electr': re.sub('[^\d\.]','',elek),
                      'teplo': re.sub('[^\d\.]','',teplo),
                      'opis': opis,
                      'url': task.url,
                      'phone': re.sub('[^\d\+\,]','',phone)[:11],
                      'lico':lico,
                      'company': comp,
                      'data':re.sub('[^\d]','',data)[:4]}
		      
     
     
	  yield Task('write',project=projects,grab=grab)
     
     
     
     def task_write(self,grab,task):
	  if task.project['sub'] <> '':
	       print('*'*100)	       
	       print  task.project['sub']
	       print  task.project['punkt']
	       print  task.project['terit']
	       print  task.project['ulica']
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
	       print  task.project['data']
	      
	      
	  
	       
	       
     
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 24, task.project['adress'])
	       self.ws.write(self.result, 6, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 28, task.project['dom'])
	       self.ws.write(self.result, 7, task.project['tip'])
	       self.ws.write(self.result, 8, task.project['naz'])
	       self.ws.write(self.result, 10, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 15, task.project['ohrana'])
	       self.ws.write(self.result, 26, task.project['gaz'])
	       self.ws.write(self.result, 27, task.project['voda'])
	       self.ws.write(self.result, 29, task.project['kanaliz'])
	       self.ws.write(self.result, 34, task.project['electr'])
	       self.ws.write(self.result, 35, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'ARENDATOR.RU')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 33, task.project['company'])
	       self.ws.write(self.result, 17, task.project['data'])
	       #self.ws.write(self.result, 32, task.project['data1'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
     
	       print('*'*10)
	       #print self.sub
	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print task.project['dom']
	       print('*'*10)
	       self.result+= 1

bot = move_Com(thread_number=5, network_try_limit=2000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5000, connect_timeout=5000)
try:
     bot.run()
except KeyboardInterrupt:
     pass
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
bot.workbook.close()
print('Done...')
time.sleep(5)
os.system("/home/oleg/pars/small/gde_zem.py")
    
       
     
     
     
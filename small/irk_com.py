#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
#from grab import Grab
import re
from sub import conv
import time
import xlsxwriter
from datetime import datetime,timedelta
import sys
import os
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'comm/0001-0081_00_C_001-0068_IRK-RU.xlsx')


class Farpost_Com(Spider):
     def prepare(self):
	  #self.f = page
	  self.ws = workbook.add_worksheet(u'irk_Коммерческая')
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
	  #self.r = conv     
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(12):
               yield Task ('post',url='http://realty.irk.ru/comm/city/1/city/9/city/122/city/121/offer/sell/offer/lease/comm_nazn/2/comm_nazn/8/comm_nazn/4/comm_nazn/64/comm_nazn/-78/date/all/order_by/promo/order/asc/pageno/%d'%x,network_try_count=100)
	  
			      
     def task_post(self,grab,task):    
	  for elem in grab.doc.select(u'//a[@class="search-link"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
     
	          
        
     def task_item(self, grab, task):
	 
	  try:
	       r = grab.doc.select(u'//p[contains(text(),"Адрес")]/following::td[1]/p/text()').text()
               if u'район' in r:
	            ray = grab.doc.select(u'//div[@id="page_breadcrumbs"]/a[4]').text()
               else:
	            ray = ''
	  except DataNotFound:
	       ray = ''          
	
	       
	  try:
	       trassa = grab.doc.select(u'//p[contains(text(),"Назначение")]/following::td[1]/p/text()').text()
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//h1').text().split(': ')[1].split(', ')[0]
	  except IndexError:
	       udal = ''
          try:
               seg =  re.sub('[^\d\,\:]', u'',grab.doc.select(u'//span[contains(text(),"Телефон:")]/following-sibling::span[1]').text())
	       if ':' in seg:
		    seg= seg.split(':')[1].split(',')[0]
	       else:
		    seg = seg 
          except IndexError:
               seg = ''	       
	       
	  try:
               price = grab.doc.select(u'//p[contains(text(),"Цена")]/following::td[1]/p/text()').text()#.replace(u'a',u'р.')
	  except IndexError:
	       price = ''
	       
	  try:
	       plosh = grab.doc.select(u'//p[contains(text(),"Площадь")]/following::td[1]/p/text()').text()
	  except IndexError:
	       plosh = '' 
	  try:
	       cena_za = grab.doc.select(u'//span[@class="obj_tit"]').text()
	  except IndexError:
	       cena_za = '' 
	       
	  
	  try:
	       ohrana = grab.doc.select(u'//p[contains(text(),"Адрес")]/following::td[1]/p/text()').text()
	  except IndexError:
	       ohrana =''
	  try:
	       gaz = grab.doc.select(u'//p[contains(text(),"Дата размещения:")]').text().split(u'Обновлено: ')[1]
	  except IndexError:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//h1').text().split(': ')[1]
	  except IndexError:
	       voda =''
	  try:
	       kanal = re.sub(u'^.*(?=санузел)','', grab.doc.select(u'//*[contains(text(), "санузел")]').text())[:7].replace(u'санузел',u'есть')
	  except DataNotFound:
	       kanal =''
	  try:
	       elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
	  except DataNotFound:
	       elek =''
	  try:
	       teplo = grab.doc.select(u'//div[contains(text(),"Адрес")]/following-sibling::div/span').text().replace(u'Подробности о доме','')
	  except DataNotFound:
	       teplo =''  
		    
	  try:
	       opis = grab.doc.select(u'//div[@class="default_div"]').text().replace(u'Дополнительная информация: ','') 
	  except IndexError:
	       opis = ''
	       
	 	       
	  
	  try:
	       oper = grab.doc.select(u'//p[contains(text(),"Сделка")]/following::td[1]/p/text()').text()
	  except IndexError:
	       oper = ''
	       
	  try:
	       data= grab.doc.select(u'//p[contains(text(),"Дата размещения:")]').text().split(u'Дата размещения: ')[1][:10]
	    #print data
	  except IndexError:
	       data = ''
		    
	  
						   
	       
	  projects = {'url': task.url,
                      'rayon': ray,
                      'trassa': trassa,
                      'udal': udal,
	              'segment': seg,
                      'cena': price,
                      'plosh':plosh,
	              'cena_za': cena_za,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'operazia':oper,
                      'data':data }
	  try:
	       ad= grab.doc.select(u'//input[@id="ymaps_location"]').attr('value')
	       link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ad
	       yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=100)
	  except IndexError:
	       yield Task('adres',grab=grab,project=projects)	  
	  
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
	  
	  print  task.project['rayon']
	  print  task.proj['punkt']
	  print  task.proj['teritor']
	  print  task.proj['ulica']
	  print  task.proj['dom']
	  print  task.project['trassa']
	  print  task.project['udal']
	  print  task.project['segment']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['cena_za']
	  print  task.project['ohrana']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanaliz']
	  print  task.project['electr']
	  print  task.project['opis']
	  print task.project['url']
	  #print  self.phone
	  print  task.project['data']
          print  task.project['teplo']
	  
	  #global result
	  self.ws.write(self.result, 0, u'Иркутская область')
	  self.ws.write(self.result, 1, task.project['rayon'])
	  self.ws.write(self.result, 2, task.proj['punkt'])
	  self.ws.write(self.result, 3, task.proj['teritor'])
	  self.ws.write(self.result, 4, task.proj['ulica'])
	  self.ws.write(self.result, 5, task.proj['dom'])
	  self.ws.write(self.result, 9, task.project['trassa'])
	  self.ws.write(self.result, 8, task.project['udal'])
	  self.ws.write(self.result, 21 , task.project['segment'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  self.ws.write(self.result, 24, task.project['ohrana'])
	  self.ws.write(self.result, 30, task.project['gaz'])
	  self.ws.write(self.result, 33, task.project['voda'])
	  #self.ws.write(self.result, 22, self.lico)
	  self.ws.write(self.result, 23, task.project['cena_za'])
	  self.ws.write(self.result, 19, u'IRK.RU')
	  self.ws.write_string(self.result, 20, task.project['url'])
	  self.ws.write(self.result, 18, task.project['opis'])
	  self.ws.write(self.result, 29, task.project['data'])
	  self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 28, task.project['operazia'])
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  print  task.project['operazia']
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  #if self.result > 10:
	       #self.stop()

     
bot = Farpost_Com(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
command = 'mount -a'
os.system('echo %s|sudo -S %s' % ('1122', command))
time.sleep(2)
workbook.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/small/olmp_zem.py")
 








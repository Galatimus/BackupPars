#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
#from grab import Grab
import logging
import re
import time
import os
import xlsxwriter
from datetime import datetime
import sys
reload(sys)
sys.setdefaultencoding('utf-8')




logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)





class roszem(Spider):
     def prepare(self):
	  #self.f = page
	  #self.link =l[i]
	  self.workbook = xlsxwriter.Workbook(u'zemm/0001-0039_00_У_001-0071_ROSZEM.xlsx')
	  self.ws = self.workbook.add_worksheet(u'roszem')
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
	  self.ws.write(0, 30, u"ДОРОГА")
	  self.ws.write(0, 31, u"ВИД_ПРАВА")
	  self.ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
	  
	  self.result= 1 
	  
	  
     def task_generator(self):
	  for x in range(1,371):#725
	       yield Task ('post',url='http://www.roszem.ru/search?page=%d'%x+'&sort=date_sort&type=Land',refresh_cache=True,network_try_count=100)
	  
     
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//a[@class="cl_land location"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur	      
	       yield Task('item',url=ur,network_try_count=100)
	      
               
     def task_item(self, grab, task):
	  
	 
	  try:
               sub = grab.doc.select(u'//nav[@class="wrapper"]/a[3]').text().replace(u'г. ','')
	  except IndexError:
	       sub=''
          try:
               ray = grab.doc.select(u'//nav[@class="wrapper"]/a[contains(text(),"район")]').text()
          except IndexError:
               ray = ''
	       
          try:
               ra = grab.doc.select(u'//p[@class="location"]').text()
	       punkt = ra.split(', ')[len(ra.split(','))-1]
               if ra.find(u'шоссе')>=0:
                    trassa = ra.split(', ')[0].split(' (')[0]
                    ter=''
	       else:
                    ter = ra.split(', ')[0]
                    trassa=''
               i=0
               for w in ra.split(','):
                    i+=1
                    if w.find(u'км от города')>=0:
                         udal = ra.split(', ')[i-1].replace(u' города','')
                         break
               if w.find(u'км от города')<0:
		    udal =''
          except IndexError:
               ra = ''
	       ter=''
	       trassa=''
	       udal =''
	       punkt=''
          try:
               price = grab.doc.select(u'//p[@class="price"]').text()
          except IndexError:
               price = ''
          try:
               price_sot = grab.doc.select(u'//th[@class="price_per"][contains(text(),"Цена за сотку")]/following::p[@class="price"][2]').text()
          except IndexError:
               price_sot = ''
          try:
               plosh = grab.doc.select(u'//p[@class="square t-center"]').text()
          except IndexError:
               plosh = ''
          try:
               kat = grab.doc.select(u'//dt[contains(text(),"Категория земель:")]/following-sibling::dd').text()
          except IndexError:
               kat = ''
          try:
               vid = grab.doc.select(u'//dt[contains(text(),"Вид разрешенного использования:")]/following-sibling::dd').text()
          except IndexError:
               vid = ''
          try:
	       gaz = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"газ")]').text().replace(u'есть газ',u'есть').replace(u'нет газа','')
	  except IndexError:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"водоснабжение")]').text().replace(u'есть водоснабжение',u'есть').replace(u'нет водоснабжения','')
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"канализация")]').text().replace(u'есть канализация',u'есть').replace(u'нет канализации','')
	  except IndexError:
	       kanal =''
	  try:
	       elekt = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"электричество")]').text().replace(u'есть электричество',u'есть').replace(u'нет электричества','')
	  except IndexError:
	       elekt =''
          try:
               teplo = grab.doc.select(u'//p[@class="location"]').text()
          except IndexError:
               teplo =''
          try:
               ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
          except IndexError:
               ohrana =''
          try:
               opis = grab.doc.select(u'//h3[contains(text(),"Описание")]/following-sibling::p').text() 
          except IndexError:
               opis = ''
          try:
               phone = re.sub('[^\d\+\,]','',grab.doc.select(u'//dt[contains(text(),"Телефон")]/following-sibling::dd').text())
          except IndexError:
               phone = ''
          try:
               lico = grab.doc.select(u'//dt[contains(text(),"Продавец")]/following-sibling::dd').text()
          except IndexError:
               lico = ''
          try:
               data = re.sub(u'^.*Размещен ','',grab.doc.select(u'//p[@class="id"]').text()).replace(')','')
          except IndexError:
               data = ''
          try:
               doroga = grab.doc.select(u'//dt[contains(text(),"Транспортная доступность:")]/following-sibling::dd').text()
          except IndexError:
               doroga = ''
          try:
               pravo = grab.doc.select(u'//dt[contains(text(),"Вид права:")]/following-sibling::dd').text()
          except IndexError:
               pravo = ''
	       
	       
	       
                       
	  projects = {'url': task.url,
	              'sub':sub,
	              'rayon': ray,
	              'punkt': re.sub('[0-9]', u'',punkt).replace(u'км от города','').replace(u'(МКАД)','').replace(u'шоссе',''),
	              'teritory':re.sub('[0-9]', u'',ter).replace(u'км от города','').replace(u'(МКАД)',''),
	              'trassa': trassa,
	              'udal': udal,
	              'price': price,
	              'price_sot': price_sot,
	              'ploshad': plosh,
	              'kategory': kat,
	              'vid': vid,
	              'gaz': gaz,
	              'voda':voda,
	              'kanal': kanal,
	              'elekt': elekt,
	              'teplo': teplo,
	              'ohrana': ohrana,
	              'opis': opis,
	              'phone': phone,
	              'lico':lico,
	              'dataraz': data,
	              'doroga': doroga,
	              'pravo': pravo
	              }
		          
	       
	       
	  yield Task('write',project=projects,grab=grab)
          
          
          
          
          
     def task_write(self,grab,task):
	       
	  print('*'*100)
	  print  task.project['sub']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['teritory']
	  print  task.project['trassa']
	  print  task.project['udal']
	  print  task.project['price']
	  print  task.project['price_sot']
	  print  task.project['ploshad']
	  print  task.project['kategory']
	  print  task.project['vid']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanal']
	  print  task.project['elekt']	  
	  print  task.project['ohrana']
          print  task.project['opis']
          print task.project['url']
          print  task.project['phone']
          print  task.project['lico']
          print  task.project['dataraz']
	  print  task.project['teplo']
	  print  task.project['doroga']
	  print  task.project['pravo']	  
	      
	      
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 3, task.project['teritory'])
	  #self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 7, task.project['trassa'])
	  self.ws.write(self.result, 8, task.project['udal'])
	  self.ws.write(self.result, 9, u'Продажа')
	  self.ws.write(self.result, 10, task.project['price'])
	  self.ws.write(self.result, 11, task.project['price_sot'])
	  self.ws.write(self.result, 12, task.project['ploshad'])
	  self.ws.write(self.result, 13, task.project['kategory'])
	  self.ws.write(self.result, 14, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanal'])
	  self.ws.write(self.result, 18, task.project['elekt'])
	  self.ws.write(self.result, 32, task.project['teplo'])
	  self.ws.write(self.result, 20, task.project['ohrana'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'ROSZEM.RU')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  #self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['dataraz'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 30, task.project['doroga'])
	  self.ws.write(self.result, 31, task.project['pravo'])	  
	  
	 
	  print('*'*100)
	 
	  print 'Ready - '+str(self.result)
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  
	  print('*'*100)
	  self.result+= 1
	       
	  #if self.result > 50:
	       #self.stop()
          
               
               
          
          
    
    
    
bot = roszem(thread_number=3, network_try_limit=1000)
#bot.setup_queue('mongo', database='Roszem',host='192.168.10.200')
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
time.sleep(5)
os.system("/home/oleg/pars/small/upn_zem.py")


#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import time
import re
from datetime import datetime
import xlsxwriter
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

workbook = xlsxwriter.Workbook(u'0001-0199_00_У_001-0220_TORGI.xlsx')
class Torgi_Zem(Spider):
     def prepare(self):
	  self.ws = workbook.add_worksheet()
	  self.ws.write(0, 0, u"Организатор торгов")
	  self.ws.write(0, 1, u"Срок подведения итогов")
	  self.ws.write(0, 2, u"Статус торгов")
	  self.ws.write(0, 3, u"Победитель торгов")
	  self.ws.write(0, 4, u"Предмет торга")
	  self.ws.write(0, 5, u"Цена, предложенная победителем")
	  self.ws.write(0, 6, u"Результат торгов")
	  self.ws.write(0, 7, u"Тип торгов")
	  self.ws.write(0, 8, u"Вид собственности")
	  self.ws.write(0, 9, u"Кадастровый номер")
	  self.ws.write(0, 10, u"Целевое назначение и разрешенное использование земельного участка")
	  self.ws.write(0, 11, u"Местоположение имущества")
	  self.ws.write(0, 12, u"Детальное местоположение")
	  self.ws.write(0, 13, u"Площадь/м2")
	  self.ws.write(0, 14, u"Срок аренды")
	  self.ws.write(0, 15, u"Валюта лота")
	  self.ws.write(0, 16, u"Ежемесячный платеж за объект")
	  self.ws.write(0, 17, u"Ежегодный платеж за объект")
	  self.ws.write(0, 18, u"Обременение")
	  self.ws.write(0, 19, u"Описание обременения")
	  self.ws.write(0, 20, u"Дата, время и порядок осмотра земельного участка на местности")
	  self.ws.write(0, 21, u"Порядок определения победителей торгов")
	  self.ws.write(0, 22, u"URL")
	  self.ws.write(0, 23, u"Документы результатов")
	  self.ws.write(0, 24, u"Основания размещения извещения")
	  self.ws.write(0, 25, u"Ежемесячная начальная цена/м2")
	  self.ws.write(0, 26, u"Дата окончания приема заявок")
	  self.ws.write(0, 27, u"Описание")
	  self.ws.write(0, 28, u"Дата_парсинга")
	  self.ws.write(0, 29, u"Дата_отмены")
	  self.ws.write(0, 30, u"Дата и время проведения аукциона")
	  self.ws.write(0, 31, u"Дата окончания приема заявок")
	  self.ws.write(0, 32, u"Начальная цена")
	  self.ws.write(0, 33, u"Вид разрешенного использования")
	  self.ws.write(0, 34, u"Категория земель")
	  self.ws.write(0, 35, u"Дата и время публикации извещения")
	  self.ws.write(0, 36, u"Телефон")
	  self.ws.write(0, 37, u"Цель предоставления земельного участка")
	  self.ws.write(0, 38, u"Минимальная цена")
	  self.result= 1
     def task_generator(self):
	  l= open('links/Torgi_Zem.txt').read().splitlines()
	  self.dc = len(l)
	  print self.dc
	  for line in l:
	       yield Task ('item',url=line.strip(),refresh_cache=True,network_try_count=100)
        
        
     def task_item(self, grab, task):
	  try:
	       sub= grab.doc.select(u'//label[contains(text(),"Организатор торгов")]/following::span[1]').text()
	  except DataNotFound:
	       sub=''
	  try:
	       srok= grab.doc.select(u'//label[contains(text(),"Срок подведения итогов")]/following::span[1]').text()
	  except DataNotFound:
	       srok=''
	  try:
	       ray = grab.doc.select(u'//label[contains(text(),"Дата и время публикации извещения")]/following::span[1]').text()[:10]
	  except DataNotFound:
	       ray = ''
	  try:
	       punkt = grab.doc.select(u'//label[contains(text(),"Дата и время окончания приема заявок")]/following::span[1]').text()[:10]
	  except DataNotFound:
	       punkt = ''
	  try:
	       rezi = grab.doc.select(u'//label[contains(text(),"Результат торгов:")]/following::span[1]').text()
	  except DataNotFound:
	       rezi = ''	       
	 
	  try:
	       dom = grab.doc.select(u'//span[@id="article"]').text()
	  except IndexError:
	       dom = ''
	  try:
	       trassa = grab.doc.select(u'//span[@id="startPrice"]').text()
	       
	  except IndexError:
	       trassa = ''
	  try:
	       udal = grab.doc.select(u'//span[@id="bidType"]').text()
	      
	  except IndexError:
	       udal = ''
	  try:
	       
	       price = grab.doc.select('//span[@id="propKind"]').text()
	       
	  except IndexError:
	       price = ''
	  try:
	       
	       price_sot = grab.doc.select('//span[@id="reqDecision"]').text()
	       
	  except IndexError:
	       price_sot = ''
	  try:
	       plosh = grab.doc.select(u'//span[@id="cadastralNum"]').text()
	  except IndexError:
	       plosh = ''
	  try:
	       vid = grab.doc.select(u'//span[@id="mission"]').text()
	  except IndexError:
	       vid = ''
	  try:
	       gaz = grab.doc.select(u'//span[@id="readonlyKladr"]').text()
	  except IndexError:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//span[@id="location"]').text()
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//label[contains(text(),"Площадь м²:")]/following::span[1]').text()
	  except IndexError:
	       kanal =''
	  try:
	       elekt = grab.doc.select(u'//span[@id="description"]').text()
	  except IndexError:
	       elekt =''
	  try:
	       teplo = grab.doc.select(u'//label[contains(text(),"Срок аренды:")]/following::span[1]').text()
	  except IndexError:
	       teplo =''
	  try:
	       ohrana = grab.doc.select(u'//span[@id="currency"]').text()
	  except IndexError:
	       ohrana =''
	  try:
	       opis = grab.doc.select(u'//span[@id="step"]').text() 
	  except IndexError:
	       opis = ''
	  try:
	       ph = grab.doc.select(u'//span[@id="maintenanceSize"]').text()
	       
	  except IndexError:
	       ph = ''
       
	  try:
	       lico = grab.doc.select(u'//span[@id="depositSize"]').text()
	  except IndexError:
	       lico = ''
       
	  try:
	       com = grab.doc.select(u'//span[@id="depositDesc"]').text()
	  except IndexError:
	       com = ''
	  try:
	       data = grab.doc.select(u'//span[@id="groundViewPlace"]').text()
	  except IndexError:
	       data = ''
       
	  try:
	       data1 =  grab.doc.select(u'//label[contains(text(),"Статус лота:")]/following::span[1]').text()
	  except IndexError:
	       data1 = ''

	  try:
	       data2 =  grab.doc.select(u'//span[@id="techConditions"]').text()
	  except IndexError:
	       data2 = ''
	       
	  try:
	       win =  grab.doc.select(u'//label[contains(text(),"Единственный участник:")]/following::span[1]').text()
	  except IndexError:
	       win = ''
	  try:
	       osnova =  grab.doc.select(u'//label[contains(text(),"Основание размещения извещения:")]/following::span[1]').text()
	  except IndexError:
	       osnova = ''
	  try:
	       nach_cena =  grab.doc.select(u'//label[contains(text(),"Ежемесячная начальная цена 1 кв.м:")]/following::span[1]').text()
	  except IndexError:
	       nach_cena = ''
	  try:
	       plat =  grab.doc.select(u'//label[contains(text(),"Ежемесячный платеж за объект:")]/following::span[1]').text()
	  except IndexError:
	       plat = '' 
          try:
               data_ot =  grab.doc.select(u'//label[contains(text(),"Дата отмены:")]/following::span[1]/p').text()
          except IndexError:
	       data_ot = '' 
	  try:
	       vid_isp =  grab.doc.select(u'//label[contains(text(),"Вид разрешенного использования:")]/following::span[1]').text()
	  except IndexError:
	       vid_isp = ''
	  try:
	       cat =  grab.doc.select(u'//label[contains(text(),"Категория земель:")]/following::span[1]').text()
	  except IndexError:
	       cat = ''	
	  try:
	       cel =  grab.doc.select(u'//label[contains(text(),"Цель предоставления земельного участка:")]/following::span[1]').text()
          except IndexError:
	       cel = ''	      
	       
	  try:
	       min_cena =  grab.doc.select(u'//label[contains(text(),"Минимальная цена")]/following::span[1]').text()
          except IndexError:
	       min_cena = ''	       
	       

   
	  projects = {'url': task.url,
                      'sub': sub,
	              'srok': srok,
                      'rayon': ray,
                      'punkt': punkt,
	              'rezult': rezi,
                      'dom': dom,
                      'trassa': trassa,
                      'udal': udal,
                      'price': price,
                      'price_sot': price_sot,
                      'ploshad': plosh,
                      'vid': vid,
                      'gaz': gaz,
                      'voda':voda,
                      'kanal': kanal,
                      'elekt': elekt,
                      'teplo': teplo,
                      'ohrana': ohrana,
                      'opis': opis,
	              'oplata': plat,
                      'phone': ph,
                      'lico':lico,
                      'company':com,
                      'win':win,
	              'osnovan':osnova,
	              'cel':cel,
	              'cena1':nach_cena,
                      'dataraz': data,
                      'data1': data1,
                      'data2': data2,
	              'cenam': min_cena,
	              'dataot':data_ot,
	              'vid_ispol':vid_isp,
	              'category':cat}
                      ####################
                      #'itogi': itogi}
	  gr = grab.clone()
	  gr.setup(url='https://torgi.gov.ru/?wicket:interface=:0:notificationEditPanel:tabs:tabs-container-parent:tabs-container:tabs:0:link::IBehaviorListener:0:2')
	  yield Task('next',grab=gr,project=projects,refresh_cache=True,network_try_count=100)
	  del gr
	  
	  
     def task_next(self, grab, task):
	  try:
	       itogi = grab.doc.select(u'//label[contains(text(),"Дата и время проведения аукциона:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       itogi = ''	  
	  
	  try:
	       phone =  grab.doc.select(u'//label[contains(text(),"Телефон:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       phone = ''  
	  
          project2 ={'itogi': itogi,
	             'phone':phone}
          
	  yield Task('write',project=task.project,proj=project2,grab=grab)
	

	
	
	
	
     def task_write(self,grab,task):
	
	  print('*'*100)
	  print  task.project['sub']
	  print  task.project['srok']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['rezult']
	  print  task.project['dom']
	  print  task.project['trassa']
	  print  task.project['udal']
	  print  task.project['price']
	  print  task.project['price_sot']
	  print  task.project['ploshad']
	  print  task.project['oplata']
	  print  task.project['vid']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanal']
	  print  task.project['elekt']
	  print  task.project['teplo']
	  print  task.project['ohrana']
	  print  task.project['opis']
	  print task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['dataraz']
	  print  task.project['data1']
	  print  task.project['data2']
	  print  task.project['win']
	  print  task.project['osnovan']
	  print  task.project['cena1']
	  print  task.project['dataot']
	  print  task.proj['itogi']
	  print  task.project['vid_ispol']
	  print  task.project['category']
	  print  task.proj['phone']
	  print  task.project['cenam']
	  
	  self.ws.write(self.result, 27,u'Организатор торгов - '+ task.project['sub']+u' , Дата публикации извещения - '
	                +task.project['rayon']+u' , Дата окончания приема заявок - '
	                +task.project['punkt']+u' , Начальная цена - '
	                +task.project['trassa']+u' , Реквизиты решения о проведении торгов - '
	                +task.project['price_sot']+u' , Описание границ земельного участка - '
	                +task.project['elekt']+u' , Коммуникации - '
	                +task.project['data2']+u' , Шаг аукциона - '
	                +task.project['opis']+u' , Размер обеспечения - '
	                +task.project['phone']+u' , Размер задатка - '
	                +task.project['lico']+u' , Порядок внесения и возврата задатка - '
	                +task.project['company'])
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['srok'])
	  self.ws.write(self.result, 2, task.project['data1'])
	  self.ws.write(self.result, 3, task.project['win'])
	  self.ws.write(self.result, 4, task.project['dom'])
	  self.ws.write(self.result, 6, task.project['rezult'])
	  self.ws.write(self.result, 7, task.project['udal'])
	  self.ws.write(self.result, 8, task.project['price'])
	  self.ws.write(self.result, 9, task.project['ploshad'])
	  self.ws.write(self.result, 10, task.project['vid'])
	  self.ws.write(self.result, 11, task.project['gaz'])
	  self.ws.write(self.result, 12, task.project['voda'])
	  self.ws.write(self.result, 13, task.project['kanal'])
	  self.ws.write(self.result, 14, task.project['teplo'])
	  self.ws.write(self.result, 15, task.project['ohrana'])
	  self.ws.write(self.result, 16, task.project['oplata'])
	  self.ws.write(self.result, 24, task.project['osnovan'])
	  self.ws.write(self.result, 36, task.proj['phone'])
	  self.ws.write(self.result, 25, task.project['cena1'])
	  self.ws.write(self.result, 20, task.project['dataraz'])
	  self.ws.write(self.result, 21,u'Победителем аукциона признается участник, предложивший наивысшую цену в ходе проведения торгов')
	  self.ws.write_string(self.result, 22, task.project['url'])
	  self.ws.write(self.result, 28, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 29, task.project['dataot'])
	  self.ws.write(self.result, 30, task.proj['itogi'])
	  self.ws.write(self.result, 31, task.project['punkt'])
	  self.ws.write(self.result, 32, task.project['trassa'])
	  self.ws.write(self.result, 33, task.project['vid_ispol'])
	  self.ws.write(self.result, 34, task.project['category'])
	  self.ws.write(self.result, 35, task.project['rayon'])
	  self.ws.write(self.result, 37, task.project['cel'])
	  self.ws.write(self.result, 38, task.project['cenam'])
	  
   
	  print('*'*100)
	  print 'Ready - '+str(self.result)+'/'+str(self.dc)
	  logger.debug('Tasks - %s' % self.task_queue.size()) 
	  print('*'*100)
	  
	  
	  self.result+= 1
	  
	  
	  #if self.result > 50:
	       #self.stop()
	

bot = Torgi_Zem(thread_number=15, network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=500, connect_timeout=500)
try:
     bot.run()
except KeyboardInterrupt:
     pass
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')    
command = 'mount -a'
p = os.system('echo %s|sudo -S %s' % ('1122', command))
print p
time.sleep(2)
workbook.close()
print('Done!')

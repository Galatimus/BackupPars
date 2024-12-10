#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from datetime import datetime,timedelta
import time
import re
import os
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


logging.basicConfig(level=logging.DEBUG)





workbook = xlsxwriter.Workbook(u'zemm/0001-0013_00_У_001-0010_ZEM-RU.xlsx')


class Zemru_Com(Spider):


     def prepare(self):
	  self.ws = workbook.add_worksheet(u'ZemRu_Земля')
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
	  self.result= 1
	  #self.num = 1
	  
     
     
     
     
     
     def task_generator(self):
	  for x in range(1,129):
	       yield Task ('post',url='http://base.zem.ru/all/?page=%d'%x+'&cat=1',network_try_count=100)
          for x1 in range(1, 178):
               yield Task ('post',url='http://base.zem.ru/all/?page=%d'%x1+'&cat=3',network_try_count=100) 
          for x2 in range(1, 13):
               yield Task ('post',url='http://base.zem.ru/all/?page=%d'%x2+'&cat=9',network_try_count=100)
       
     
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//a[@class="b-special-lots__item__information__link-title-text"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur	      
	       yield Task('item',url=ur,refresh_cache=True,network_try_count=100)
	 
               
               
               
               
     def task_item(self, grab, task):
	  try:
	       
	       sub = grab.doc.select(u'//span[contains(text(),"Область/Край")]/following-sibling::span').text()
	  except IndexError:
	       sub =''
          try:
               ray =  grab.doc.select(u'//span[contains(text(),"Регион")]/following-sibling::span[contains(text(),"район")]').text()
          except IndexError:
	       ray=''
          try:
               punkt = grab.doc.select(u'//span[contains(text(),"Населённый пункт")]/following-sibling::span').text()
          except IndexError:
               punkt = ''
	  try:
	       oper = grab.doc.select('//span[@class="b-title__lot__sale"]').text()
	  except IndexError:
	       oper = ''
	  try:
	       price = grab.doc.select(u'//span[@class="b-sidebar-info__object__cost__value"]').text()
          except IndexError:
               price = ''
          try:
               plosh = grab.doc.select(u'//span[contains(text(),"Площадь участка")]/following-sibling::span').text()
          except IndexError:
	       plosh = ''
          try:
               cat = grab.doc.select(u'//span[contains(text(),"Категория земель")]/following-sibling::span').text()
          except IndexError:
               cat = '' 
          try:
               vid = grab.doc.select(u'//span[contains(text(),"Подкатегория")]/following-sibling::span').text()
          except IndexError:
               vid = ''
          try:
               gaz = grab.doc.select(u'//span[contains(text(),"Газоснобжение")]/following-sibling::span').text()
          except IndexError:
               gaz =''
          try:
               voda = grab.doc.select(u'//span[contains(text(),"Водоснабжение")]/following-sibling::span').text()
          except IndexError:
               voda =''
          try:
               kanal = grab.doc.select(u'//span[contains(text(),"Канализация")]/following-sibling::span').text()
          except IndexError:
               kanal =''
          try:
               elek =  grab.doc.select(u'//span[contains(text(),"Электроснабжение")]/following-sibling::span').text()
          except IndexError:
               elek =''
          try:
               teplo = grab.doc.select(u'//span[contains(text(),"Отопление")]/following-sibling::span').text()
          except IndexError:
               teplo =''
          try:
               opis = grab.doc.select(u'//div[@class="b-object__description__text"]').text() 
          except IndexError:
               opis = ''
          try:
               phone = re.sub('[^\d\+]', '',grab.doc.select(u'//span[contains(text(),"Телефон:")]/following-sibling::span').text())
          except IndexError:
               phone = ''
          try:
               lico = grab.doc.select(u'//span[contains(text(),"Лот создал")]/following-sibling::span').text()
          except IndexError:
               lico = ''
          try:
               comp = grab.doc.select(u'//a[@class="b-sidebar-vendor-link"]').text()
          except IndexError:
               comp = ''	       
	       
	  try:
	       data = grab.doc.select(u'//span[contains(text(),"Дата публикации")]/following-sibling::span').text()
	  except IndexError:
	       data = ''
	  
	       
	       
	       
	       
	       
          
                       
	  projects = {'url': task.url,
	              'sub': sub,
	              'rayon': ray,
	              'punkt': punkt,
	              'cena': price,
	              'plosh':plosh,
	              'vid': vid,
	              'cat':cat,
	              'gaz': gaz,
	              'voda': voda,
	              'kanaliz': kanal,
	              'electr': elek,
	              'teplo': teplo,
	              'opis':opis,
	              'phone':phone,
	              'lico':lico,
	              'company':comp,
	              'data':data,
	              'oper':oper
	                 }
     
     
	  yield Task('write',project=projects,grab=grab)
          
          
          
          
          
     def task_write(self,grab,task):
	  
	       
	  print('*'*50)
	  print  task.project['sub']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['vid']
	  print  task.project['cat']
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
	  print  task.project['data']
	  print  task.project['oper']
	  
	  
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 14, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 19, task.project['teplo'])
	  self.ws.write(self.result, 13, task.project['cat'])	       
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'ZEM.RU')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	  
	      
	  print('*'*50)
	  print 'Ready - '+str(self.result)
	  print 'Tasks - %s' % self.task_queue.size()
	  print('*'*50)	       
	  self.result+= 1
	  
	  
	  
	  
	  #if self.result > 100:
               #self.stop()	  
	  
          
               
               
          
          
      
    
    
bot = Zemru_Com(thread_number=5, network_try_limit=1000)
#bot.setup_queue('mongo', database='ZemRu',host='192.168.10.200')
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
command = 'mount -a'
os.system('echo %s|sudo -S %s' % ('1122', command))
time.sleep(2)
workbook.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/small/doska_zem.py")



     
     

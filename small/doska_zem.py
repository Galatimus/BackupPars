#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
#from grab import Grab
import re
import random
import time
import xlsxwriter
import time
import os
from datetime import datetime,timedelta

logging.basicConfig(level=logging.DEBUG)


workbook = xlsxwriter.Workbook(u'zemm/0001-0013_00_У_001-0023_DOSKA.xlsx')

   

class Brsn_Zem(Spider):
     def prepare(self):
	  #self.f = page
	  self.ws = workbook.add_worksheet(u'Doska_Земля')
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
	  self.ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
	  self.ws.write(0, 30, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  #for line in open('/home/oleg/mirkvartir/Links/Zem.txt').read().decode('cp1251').splitlines():
	  yield Task ('sub',url='http://www.doska.ru/real-estate/zemlja-i-u4astki/',network_try_count=100)
	  
	  
     def task_sub(self,grab,task):
	  for el in grab.doc.select(u'//h4[@class="category"]/a'):
	       urr = grab.make_url_absolute(el.attr('href'))  
	       #print urr
	       yield Task ('post',url=urr,network_try_count=100)
	  
	  
            
            
     def task_page(self,grab,task):
	  try:         
	       pg = grab.doc.select(u'//button[@class="navia"]/following-sibling::a[contains(@href,"page")][1]')
	       u = grab.make_url_absolute(pg.attr('href'))
	       yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	  except DataNotFound:
	       print('*'*50)
	       print '!!!','NO PAGE NEXT','!!!'
	       print('*'*50)
	       print '%s taskq size' % self.task_queue.size()             
        
        
            
            
     def task_post(self,grab,task):
    
	  for elem in grab.doc.select(u'//div[@class="d1"]/a'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	  yield Task("page", grab=grab,refresh_cache=True,network_try_count=100,use_proxylist=False)
        
        
     def task_item(self, grab, task):
	  try:
	       sub = grab.doc.select(u'//td[contains(text(),"Область:")]/following-sibling::td').text().split(', ')[0]
	  except IndexError:
	       sub = ''
	  try:
	       try:
	            ray = grab.doc.select(u'//td[contains(text(),"айон")]/following-sibling::td/b[contains(text(),"район")]').text()
	       except IndexError:
	            ray = grab.doc.select(u'//td[contains(text(),"айон")]/following-sibling::td[contains(text(),"район")]').text()
	  except IndexError:
	       ray = ''          
	  try:
	       if sub == u'Москва':
                    punkt= u'Москва'
               elif sub == u'Санкт-Петербург':
	            punkt= u'Санкт-Петербург'
               elif sub == u'Севастополь':
	            punkt= u'Севастополь'
	       else:
		    if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[2][contains(text(),"район")]').exists()==True:
			 punkt= grab.doc.select(u'//h1[@class="object_descr_addr"]/a[3]').text()
		    elif grab.doc.select(u'//h1[@class="object_descr_addr"]/a[3][contains(text(),"район")]').exists()==True:
			 punkt= grab.doc.select(u'//h1[@class="object_descr_addr"]/a[2]').text()
		    else:
			 punkt=grab.doc.select(u'//td[contains(text(),"Город")]/following-sibling::td').text().replace(ray,'')
	  except IndexError:
	       punkt = ''
	       
	  try:
	       ter= grab.doc.select(u'//td[contains(text(),"Район")]/following-sibling::td').text().replace(ray,'')
	  except IndexError:
	       ter =''
	       
	  try:
	       try:
	            uliza = re.split(r'(\W+)',grab.doc.select(u'//td[contains(text(),"Адрес")]/following-sibling::td').text(),1)[1].replace(u' [Карта]','')
	       except IndexError:
	            uliza =  re.split(r'(\W+)',grab.doc.select(u'//td[contains(text(),"Улица")]/following-sibling::td').text(),1)[1].replace(u' [Карта]','')
	  except IndexError:
	       uliza = ''
	       
          try:
               dom = re.split('\W+', uliza,1)[1]
          except (IndexError,AttributeError):
               dom = ''
	       
	  try:
	       trassa = grab.doc.select(u'//td[contains(text(),"Цена:")]/following-sibling::td').text().split(' (')[1].replace(')','')
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//label[contains(text(),"Шоссе:")]/following-sibling::p').text().split(', ')[1]
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//td[contains(text(),"Цена:")]/following-sibling::td').text().split(' (')[0]
	  except IndexError:
	       price = ''

	  try:
	       plosh = grab.doc.select(u'//td[contains(text(),"Площадь:")]/following-sibling::td').text()
	  except DataNotFound:
	       plosh = ''
	  try:
	       vid = grab.doc.select(u'//td[contains(text(),"Назначение:")]/following-sibling::td').text()
	  except DataNotFound:
	       vid = '' 
	  
	  ohrana =''
	  try:
	       ad=[]
               for s in grab.doc.select(u'//td[@class="ads_opt_name"]/following-sibling::td/b'):#+str(random.randint(1,99)))#.capitalize()#.split(u' в ')[0])#.split(' ')[1].replace(u'квартир',u'Квартира').replace(u'комнат',u'Комната')
                    ad.append(s.text())
               gaz = ','.join(ad)
	  except DataNotFound:
	       gaz =''
	  try:
	       voda = re.sub(u'^.*(?=вод)','', grab.doc.select(u'//*[contains(text(), "вод")]').text())[:3].replace(u'вод',u'есть')
	  except DataNotFound:
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
	       teplo =  grab.doc.select(u'//div[@class="sttnmls_navigation text-center"]').text().replace(u' → ',', ')
	  except DataNotFound:
	       teplo =''
	       
	  try:
	       if u'/мес' in price:
		    oper = u'Аренда'
	       else:
		    oper =u'Продажа'
	  except DataNotFound:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//div[@id="msg_div_msg"]').text() 
	  except DataNotFound:
	       opis = ''
	       
	  try:
	       phone = grab.doc.select(u'//span[@id="phone_td_1"]').text().replace('***',str(random.randint(101,999)))
	  except IndexError:
	       phone = ''
	       
	  try:
	       lico = grab.doc.select(u'//img[@class="thumbnail"][contains(@src,"components")]/following::b[1]').text()
	  except IndexError:
	       lico = ''
	       
	  try:
	       comp = grab.doc.select(u'//img[@class="thumbnail"][contains(@src,"agency")]/following::b[1]').text()
	  except IndexError:
	       comp = ''
	       
	  try:
	       data= grab.doc.select(u'//td[@class="msg_footer"][contains(text(),"Дата:")]').text().split(': ')[1].split(' ')[0]
	    #print data
	  except IndexError:
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
	  print  task.project['teplo']
	  print  task.project['opis']
	  print task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	  print  task.project['oper']
	  
	  #global result
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 3, task.project['teritor'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 11, task.project['trassa'])
	  self.ws.write(self.result, 8, task.project['udal'])
	  self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write_string(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 14, task.project['vid'])
	  self.ws.write(self.result, 31, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 31, task.project['teplo'])
	  self.ws.write(self.result, 28, task.project['ohrana'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'Доска.ру')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  print 'Tasks - %s' % self.task_queue.size()
	  #print '*',i+1,'/',dc,'*'
	  #print oper
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  #if self.result > 20:
	       #self.stop()

     
bot = Brsn_Zem(thread_number=3,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
try:
     command = 'mount -a'
     os.system('echo %s|sudo -S %s' % ('1122', command))
     time.sleep(3)
     workbook.close()
     print('Done')
except IOError:
     time.sleep(30)
     os.system('echo %s|sudo -S %s' % ('1122', 'mount -a'))
     time.sleep(10)
     workbook.close()
     print('Done!') 


time.sleep(5)
os.system("/home/oleg/pars/small/doska_com.py")





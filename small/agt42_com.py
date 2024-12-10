#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
#from grab import Grab
import re
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


workbook = xlsxwriter.Workbook(u'comm/0001-0016_00_C_001-0017_AGNT42.xlsx')

    

class Ya39_Com(Spider):
     def prepare(self):
	  #self.f = page
	  self.ws = workbook.add_worksheet()
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
	  self.ws.write(0, 12, u"ПЛОЩАДЬ")
	  self.ws.write(0, 13, u"ЭТАЖ")
	  self.ws.write(0, 14, u"ЭТАЖНОСТЬ")
	  self.ws.write(0, 15, u"ГОД_ПОСТРОЙКИ")
	  self.ws.write(0, 16, u"МАТЕРИАЛ_СТЕН")
	  self.ws.write(0, 17, u"ВЫСОТА_ПОТОЛКА")
	  self.ws.write(0, 18, u"СОСТОЯНИЕ")
	  self.ws.write(0, 19, u"БЕЗОПАСНОСТЬ")
	  self.ws.write(0, 20, u"ГАЗОСНАБЖЕНИЕ")
	  self.ws.write(0, 21, u"ВОДОСНАБЖЕНИЕ")
	  self.ws.write(0, 22, u"КАНАЛИЗАЦИЯ")
	  self.ws.write(0, 23, u"ЭЛЕКТРОСНАБЖЕНИЕ")
	  self.ws.write(0, 24, u"ТЕПЛОСНАБЖЕНИЕ")
	  self.ws.write(0, 25, u"ОПИСАНИЕ")
	  self.ws.write(0, 26, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 27, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 28, u"ТЕЛЕФОН")
	  self.ws.write(0, 29, u"КОНТАКТНОЕ_ЛИЦО")
	  self.ws.write(0, 30, u"КОМПАНИЯ")
	  self.ws.write(0, 31, u"ДАТА_РАЗМЕЩЕНИЯ")
	  self.ws.write(0, 32, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 33, u"ОПЕРАЦИЯ")
	  self.ws.write(0, 34, u"МЕСТОПОЛОЖЕНИЕ")
	  self.ws.write(0, 35, u"ЦЕНА_M2")
	       
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(8):#30
               yield Task ('post',url='http://agent42.ru/index.php?option=com_adsmanager&page=show_category&catid=22&order=25&ordero=&expand=0&Itemid=546&limit=20&limitstart='+str(x*20),network_try_count=100)
	  for x1 in range(6):#18
               yield Task ('post',url='http://agent42.ru/index.php?option=com_adsmanager&page=show_category&catid=37&order=0&ordero=&expand=0&Itemid=546&limit=20&limitstart='+str(x1*20),network_try_count=100)
         
                 
        
        
            
            
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//h2/a'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	 
        
        
     def task_item(self, grab, task):
	  try:
	       sub = u'Кемеровская область'
	  except DataNotFound:
	       sub = ''
	  try:
	       ray = grab.doc.select(u'//div[@class="adsmanager_ads_contact"]').text().replace(u'Расположение: ','').replace(u'Адрес: ',', ').replace(u'Район: ',', ')
	     #print ray 
	  except DataNotFound:
	       ray = ''          
	  try:
	       punkt= u'Кемерово'#grab.doc.select(u'//div[contains(text(),"Адрес:")]/following-sibling::div').text().split(u', ')[0].replace(sub,'')
	  except IndexError:
	       punkt = ''
	       
	  try:
	       #if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text()," район")]').exists()==True:
	       ter= grab.doc.select(u'//b[contains(text(),"Район:")]/following-sibling::text()').text().replace(u'г. Кемерово ','')
	       #else:
		    #ter= ''#grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[2]
	  except IndexError:
	       ter =''
	       
	  try:
	       #if grab.doc.select(u'').exists()==False:
	       uliza = re.split(r'(\W+)',grab.doc.select(u'//b[contains(text(),"Адрес:")]/following-sibling::text()[1]').text().split(u' - ')[0],1)[1]
	       
	  except IndexError:
	       uliza = ''
	  try:
	       dom = re.split('\W+',grab.doc.select(u'//b[contains(text(),"Адрес:")]/following-sibling::text()[1]').text().split(u' - ')[0],1)[1]
	  except (IndexError,AttributeError):
	       dom = ''       
	  try:
	       trassa = grab.doc.select(u'//b[contains(text(),"Адрес:")]/following-sibling::text()[1]').text().split(u' - ')[1]
		#print rayon
	  except IndexError:
	       trassa = ''       
	  try:
	       udal = grab.doc.select(u'//div[@class="adsmanager_ads_kindof"]').text().split('.')[0]
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//div[@class="adsmanager_ads_price"]').text().split('Цена: ')[1].split(u' (')[0]
	  except IndexError:
	       price = ''   
	  try:
	       plosh = grab.doc.select(u'//b[contains(text(),"Площадь общая (м2):")]/following-sibling::text()[1]').text()+u' м2'
	  except IndexError:
	       plosh = ''
	  try:
	       vid = grab.doc.select(u'//div[@class="adsmanager_pathway"]/a[4]').text().split(u'недвижимость')[0]
	  except IndexError:
	       vid = '' 
	  try:
	       et = grab.doc.select(u'//div[@class="adsmanager_ads_price"]').text().split('Цена: ')[1].split(u' (')[1].replace(' )','')
	  except IndexError:
	       et = ''
	  try:
	       et2 = grab.doc.select(u'//b[contains(text(),"Высота потолков:")]/following-sibling::text()[1]').text()
	  except IndexError:
	       et2 = ''
	  
	  try:
	       mat = grab.doc.select(u'//b[contains(text(),"Материал стен:")]/following-sibling::text()[1]').text()
	  except IndexError:
	       mat = ''
          try:
               godp = grab.doc.select(u'//b[contains(text(),"Год постройки объекта")]/following-sibling::text()[1]').text()
          except IndexError:
               godp = ''	       
	       
	       
	  try:
	       ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
	  except DataNotFound:
	       ohrana =''
	  try:
	       gaz = re.sub(u'^.*(?=газ)','', grab.doc.select(u'//*[contains(text(), "газ")]').text())[:3].replace(u'газ',u'есть')
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
	       teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
	  except IndexError:
	       teplo =''
	       
	  try:
	       oper = grab.doc.select(u'//div[@class="adsmanager_pathway"]/a[2]').text().split(' недвижимости')[0].replace(u'Покупателям',u'Продажа').replace(u'Арендаторам',u'Аренда') 
	  except IndexError:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//div[@class="adsmanager_ads_desc"][2]').text().split('Контактное лицо: ')[0] 
	  except IndexError:
	       opis = ''
	       
	  try:
	       phone = re.sub('[^\d\,\;]', '',grab.doc.select(u'//b[contains(text(),"Контактное лицо:")]/following-sibling::text()[1]').text())[1:].split(';')[0]
	  except IndexError:
	       phone = ''
	       
	  try:
	       lico = re.sub(u'[^а-яА-Я\ ]','',grab.doc.select(u'//b[contains(text(),"Контактное лицо:")]/following-sibling::text()[1]').text())
	  except IndexError:
	       lico = ''
	  comp = u'Связист '
	       
	  try:
	       data= grab.doc.select(u'//b[contains(text(),"Коммуникации:")]/following-sibling::text()[1]').text()
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
	              'et': et,
	              'ets': et2,
	              'mat': mat,
	              'god':godp,
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
	  print  task.project['punkt']
	  print  task.project['teritor']
	  print  task.project['ulica']
	  print  task.project['dom']
	  print  task.project['trassa']
	  print  task.project['udal']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['et']
	  print  task.project['ets']
	  print  task.project['mat']
	  print  task.project['god']
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
	  print  task.project['rayon']
	  
	  #global result
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 34, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 3, task.project['teritor'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 6, task.project['trassa'])
	  self.ws.write(self.result, 8, task.project['udal'])
	  self.ws.write(self.result, 33, task.project['oper'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 35, task.project['et'])
	  self.ws.write(self.result, 17, task.project['ets'])
	  self.ws.write(self.result, 15, task.project['god'])
	  self.ws.write(self.result, 16, task.project['mat'])	  
	  self.ws.write(self.result, 9, task.project['vid'])
	  self.ws.write(self.result, 20, task.project['gaz'])
	  self.ws.write(self.result, 21, task.project['voda'])
	  self.ws.write(self.result, 22, task.project['kanaliz'])
	  self.ws.write(self.result, 23, task.project['electr'])
	  self.ws.write(self.result, 24, task.project['teplo'])
	  self.ws.write(self.result, 19, task.project['ohrana'])	       
	  self.ws.write(self.result, 25, task.project['opis'])
	  self.ws.write(self.result, 26, u'АН "Связист"')
	  self.ws.write_string(self.result, 27, task.project['url'])
	  self.ws.write(self.result, 28, task.project['phone'])
	  self.ws.write(self.result, 29, task.project['lico'])
	  self.ws.write(self.result, 30, task.project['company'])
	  self.ws.write(self.result, 18, task.project['data'])
	  self.ws.write(self.result, 32, datetime.today().strftime('%d.%m.%Y'))
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  #print oper
	  print('*'*50)	       
	  self.result+= 1
	  
	  
	  #if self.result > 10:
	       #self.stop()	  
	       
	 

     
bot = Ya39_Com(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5000, connect_timeout=5000)
bot.run()
workbook.close()
print('Done!')
time.sleep(5)
os.system("/home/oleg/pars/small/brsn_zem.py")








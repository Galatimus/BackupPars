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


workbook = xlsxwriter.Workbook(u'zemm/0001-0064_00_У_001-0028_DOM43.xlsx')

    

class Cian_Zem(Spider):
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
	       
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(1,52):#30
               yield Task ('post',url='http://dom43.ru/realty/property/land_area/?page=%d'%x+'&display_type=list',network_try_count=100)
          
                 
        
        
            
            
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//div[@class="property-card"]/div/a'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	 
        
        
     def task_item(self, grab, task):
	  try:
	       sub = u'Кировская область'
	  except DataNotFound:
	       sub = ''
	  try:
	       ray = re.findall(u"р-н (.*?),",grab.doc.select(u'//title').text())[0]
	  except IndexError:
	       ray = ''          
	  try:
	       try:
		    try:
	                 punkt =re.findall(u"г (.*?),",grab.doc.select(u'//title').text())[0]
	            except IndexError:
		         punkt = re.findall(u"п (.*?),",grab.doc.select(u'//title').text())[0]
	       except IndexError:
		    punkt =re.findall(u"д (.*?),",grab.doc.select(u'//title').text())[0]
	  except IndexError:
	       punkt = ''
	       
	  try:
	       try:
		    try:
	                 ter =re.findall(u"пгт (.*?),",grab.doc.select(u'//title').text())[0]
	            except IndexError:
		         ter = re.findall(u"сл (.*?),",grab.doc.select(u'//title').text())[0]
	       except IndexError:
		    ter =re.findall(u"с (.*?),",grab.doc.select(u'//title').text())[0]
	  except IndexError:
	       ter =''
	       
	  try:
	       uliza = re.findall(u"ул (.*?),",grab.doc.select(u'//title').text())[0]
	  except IndexError:
	       uliza = ''
	  dom = ''
	  udal=''     
	  try:
	       trassa = grab.doc.select(u'//h3[contains(text(),"Расположение")]/following-sibling::div/p').text()
		#print rayon
	  except DataNotFound:
	       trassa = ''
	       
	  try:
	       comp = grab.doc.select(u'//div[contains(text(),"Прoдaвeц:")]/following-sibling::div/p[1]').text().split(u', ')[1]
	  except IndexError:
	       comp = ''
	       
	  try:
	       price = re.sub('[^\d]','',grab.doc.select(u'//strong[contains(text(),"Цена:")]/following-sibling::text()').text())+u' р.'
	  except DataNotFound:
	       price = ''
	       
	  
	       
	  try:
	       plosh = grab.doc.select(u'//strong[contains(text(),"Общая площадь:")]/following-sibling::text()').text()
	  except DataNotFound:
	       plosh = ''

	  try:
	       vid = grab.doc.select(u'//strong[contains(text(),"Категория земель:")]/following-sibling::text()').text()
	  except DataNotFound:
	       vid = '' 
	       
	       
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
	       oper = u'Продажа' 
	  except IndexError:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//div[@class="realty__block realty__description"]').text() 
	  except IndexError:
	       opis = ''
	       
	  try:
	       phone = re.sub('[^\d\,]', '',grab.doc.select(u'//div[@id="phone-number"]').text())
	  except IndexError:
	       phone = ''
	       
	  try:
	       lico = grab.doc.select(u'//div[contains(text(),"Прoдaвeц:")]/following-sibling::div/p[1]').text().split(u', ')[0]
	  except IndexError:
	       lico = ''
	       
	       
	  #try:
	       #conv = [(u' августа',u'.08.'), (u' июля',u'.07.2016'),
	       #(u' мая',u'.05.2016'),(u' июня',u'.06.2016'),
	       #(u' марта',u'.03.2016'),(u' апреля',u'.04.2016'),
	       #(u' января',u'.01.2016'),(u' декабря',u'.12.2015'),
	       #(u' сентября',u'.09.2016'),(u' ноября',u'.11.2015'),
	       #(u' февраля',u'.02.2016'),(u' октября',u'.10.2016'),
	       #(u'Сегодня', (datetime.today().strftime('%d.%m.%Y'))),
	       #(u'Вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]
	       #d = grab.doc.select(u'//span[@class="glyphicon glyphicon-time"]/following-sibling::text()').text()#.split(u' добавлено ')[1]
	       #data = reduce(lambda d, r: d.replace(r[0], r[1]), conv, d)
	  #except IndexError:
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
                      'phone':phone[::-1],
                      'lico':lico,
                      'company':comp,
                      'data':data,
                      'oper':oper
                      }
          
	  yield Task('write',project=projects,grab=grab,refresh_cache=True)
            
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
	  self.ws.write(self.result, 30, task.project['trassa'])
	  self.ws.write(self.result, 14, task.project['udal'])
	  self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 13, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 19, task.project['teplo'])
	  self.ws.write(self.result, 20, task.project['ohrana'])	       
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'Недвижимость Кирова')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
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
     
bot = Cian_Zem(thread_number=3,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')    
command = 'mount -a'
p = os.system('echo %s|sudo -S %s' % ('1122', command))
print p
time.sleep(5)
workbook.close()
print('Done!') 

time.sleep(5)
os.system("/home/oleg/pars/small/dom43_com.py")






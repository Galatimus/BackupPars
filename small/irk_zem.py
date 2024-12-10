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

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

#i = 0
#l= open('/home/oleg/CIAN/Links/Zem_Prod.txt').read().splitlines()
#dc = len(l)
#page = l[i] 
#oper = u'Продажа'
     
#g = Grab(timeout=100, connect_timeout=100)

workbook = xlsxwriter.Workbook(u'zemm/0001-0080_00_У_001-0068_IRK-RU.xlsx')

#result = 1
#print r

#while True:
     #print '********************************************',i+1,'/',dc,'*******************************************'
     #wb = xlwt.Workbook(encoding=('utf -8')) 
    

class Brsn_Zem(Spider):
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
	  self.ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
	  self.ws.write(0, 30, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 31, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(3):
               yield Task ('post',url='http://realty.irk.ru/zem/city/1/city/122/city/9/direction/416/direction/414/direction/417/direction/418/direction/419/direction/421/direction/420/direction/430/zem_nazn/2/zem_nazn/4/zem_nazn/8/zem_nazn/16/zem_nazn/32/date/all/order_by/promo/order/asc/pageno/%d'%x,network_try_count=100)
	       
     def task_post(self,grab,task):
    	  for elem in grab.doc.select(u'//a[@class="search-link"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	 
        
        
     def task_item(self, grab, task):
	  try:
	       sub = u'Иркутская область'#grab.doc.select(u'//label[contains(text(),"Адрес:")]/following-sibling::p').text().split(', ')[0]
	  except IndexError:
	       sub = ''
	  try:
	       r = grab.doc.select(u'//p[contains(text(),"Адрес")]/following::td[1]/p/text()').text()
	       if u'район' in r:
		    ray = grab.doc.select(u'//div[@id="page_breadcrumbs"]/a[4]').text()
	       else:
		    ray = ''
	     #print ray 
	  except IndexError:
	       ray = ''          
	  try:
	       punkt= u'Иркутск'#grab.doc.select(u'//span[@class="hp_caption"][contains(text(),"Город:")]/following-sibling::text()').text()
	  except IndexError:
	       punkt = ''
	       
	  try:
	       ter= grab.doc.select(u'//input[@id="ymaps_location"]').attr('value').split(', ')[2].replace(u'0','')
	  except IndexError:
	       ter =''
	       
	  try:
	       
	       uliza = grab.doc.select(u'//p[contains(text(),"Адрес")]/following::td[1]/p/text()').text()
	       #else:
		    #uliza = ''
	  except IndexError:
	       uliza = ''
	       
          try:
               dom = re.split('\W+', grab.doc.select(u'//div[@class="b-realty-address go-map"]').text(),1)[1]
          except IndexError:
               dom = ''
	       
	  try:
	       trassa = grab.doc.select(u'//p[contains(text(),"Дата размещения:")]').text().split(u'Обновлено: ')[1]
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//span[@class="build_price_in_table impure"]').text()
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//span[@class="build_price_in_table pure"]').text()
	  except IndexError:
	       price = ''
	       
	  
	       
	  try:
	       plosh = grab.doc.select(u'//p[contains(text(),"Площадь")]/following::td[1]/p/text()').text()
	  except IndexError:
	       plosh = ''

	  try:
	       vid = grab.doc.select(u'//p[contains(text(),"Предполагаемое использование")]/following::td[1]/p/text()').text()
	  except IndexError:
	       vid = '' 
	       
	       
	  try:
	       ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
	  except DataNotFound:
	       ohrana =''
	  try:
	       gaz = grab.doc.select(u'//div[contains(text(),"Газ")]/following-sibling::div').text()
	  except DataNotFound:
	       gaz =''
	  try:
	       voda =  grab.doc.select(u'//div[contains(text(),"Вода")]/following-sibling::div').text()
	  except DataNotFound:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//div[contains(text(),"Канализация")]/following-sibling::div').text()
	  except DataNotFound:
	       kanal =''
	  try:
	       elek = grab.doc.select(u'//div[contains(text(),"Электричество")]/following-sibling::div').text()
	  except DataNotFound:
	       elek =''
	  try:
	       teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
	  except DataNotFound:
	       teplo =''
	       
	  try:
	       oper = grab.doc.select(u'//p[contains(text(),"Вид права")]/following::td[1]/p/text()').text() 
	  except DataNotFound:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//div[@class="default_div"]').text().replace(u'Дополнительная информация: ','') 
	  except DataNotFound:
	       opis = ''
	       
	  try:
	       #try:
	            #phone = grab.doc.select(u'//b[contains(text(),"Контакт:")]/following-sibling::text()[2]').text().replace(' ','')
	       #except DataNotFound:
	       phone = re.sub('[^\d\,:]', u'',grab.doc.select(u'//span[contains(text(),"Телефон:")]/following-sibling::span[1]').text())
	       if ':' in phone:
		    phone= phone.split(':')[1].split(',')[0]
	       else:
		    phone= phone
	  except IndexError:
	       phone = ''
	       
	  try:
	       lico = grab.doc.select(u'//span[@class="obj_tit"]').text()
	  except IndexError:
	       lico = ''
	       
	  #try:
	       #comp = u'ГрадСтрой'#grab.doc.select(u'//span[@class="hp_caption"][contains(text(),"Организация:")]/following-sibling::text()').text()
	  #except IndexError:
	  comp = ''
	       
	  try:
	       data= grab.doc.select(u'//p[contains(text(),"Дата размещения:")]').text().split(u'Дата размещения: ')[1][:10]
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
	  self.ws.write(self.result, 31, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 29, task.project['trassa'])
	  self.ws.write(self.result, 11, task.project['udal'])
	  self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 14, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 19, task.project['teplo'])
	  self.ws.write(self.result, 20, task.project['ohrana'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'IRK.RU')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 27, task.project['lico'])
	  #self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
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

     
bot = Brsn_Zem(thread_number=3,network_try_limit=1000)
#bot.setup_queue(backend='mongo', database='irk',host='192.168.10.200')
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
command = 'mount -a'# cifs //192.168.1.6/d /home/oleg/pars -o username=oleg,password=1122,_netdev,rw,iocharset=utf8,file_mode=0777,dir_mode=0777' #'mount -a -O _netdev'
p = os.system('echo %s|sudo -S %s' % ('1122', command))
print p
time.sleep(5)
workbook.close()
print('Done!')

time.sleep(5)
os.system("/home/oleg/pars/small/irk_com.py")






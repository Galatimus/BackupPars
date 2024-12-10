#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import os
import math
import random
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)




workbook = xlsxwriter.Workbook(u'zagg/QP-RU_Загород.xlsx')






    
class QP_Com(Spider):
     def prepare(self):
	  self.ws = workbook.add_worksheet()
	  self.ws.write(0, 0, "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	  self.ws.write(0, 1, "МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	  self.ws.write(0, 2, "НАСЕЛЕННЫЙ_ПУНКТ")
	  self.ws.write(0, 3, "ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	  self.ws.write(0, 4, "УЛИЦА")
	  self.ws.write(0, 5, "ДОМ")
	  self.ws.write(0, 6, "ОРИЕНТИР")
	  self.ws.write(0, 7, "ТРАССА")
	  self.ws.write(0, 8, "УДАЛЕННОСТЬ")
	  self.ws.write(0, 9, "КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
	  self.ws.write(0, 10, "ТИП_ОБЪЕКТА")
	  self.ws.write(0, 11, "ОПЕРАЦИЯ")
	  self.ws.write(0, 12, "СТОИМОСТЬ")
	  self.ws.write(0, 13, "ЦЕНА_М2")
	  self.ws.write(0, 14, "ПЛОЩАДЬ_ОБЩАЯ")
	  self.ws.write(0, 15, "КОЛИЧЕСТВО_КОМНАТ")
	  self.ws.write(0, 16, "ЭТАЖНОСТЬ")
	  self.ws.write(0, 17, "МАТЕРИАЛ_СТЕН")
	  self.ws.write(0, 18, "ГОД_ПОСТРОЙКИ")
	  self.ws.write(0, 19, "ПЛОЩАДЬ_УЧАСТКА")
	  self.ws.write(0, 20, "ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	  self.ws.write(0, 21, "ГАЗОСНАБЖЕНИЕ")
	  self.ws.write(0, 22, "ВОДОСНАБЖЕНИЕ")
	  self.ws.write(0, 23, "КАНАЛИЗАЦИЯ")
	  self.ws.write(0, 24, "ЭЛЕКТРОСНАБЖЕНИЕ")
	  self.ws.write(0, 25, "ТЕПЛОСНАБЖЕНИЕ")
	  self.ws.write(0, 26, "ЛЕС")
	  self.ws.write(0, 27, "ВОДОЕМ")
	  self.ws.write(0, 28, "БЕЗОПАСНОСТЬ")
	  self.ws.write(0, 29, "ОПИСАНИЕ")
	  self.ws.write(0, 30, "ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 31, "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	  self.ws.write(0, 32, "ТЕЛЕФОН")
	  self.ws.write(0, 33, "КОНТАКТНОЕ_ЛИЦО")
	  self.ws.write(0, 34, "КОМПАНИЯ")
	  self.ws.write(0, 35, "ДАТА_РАЗМЕЩЕНИЯ")
	  self.ws.write(0, 36, "ДАТА_ОБНОВЛЕНИЯ")
	  self.ws.write(0, 37, "ДАТА_ПАРСИНГА")
	  self.ws.write(0, 38, "ЗАГОЛОВОК")
	  self.ws.write(0, 39, "МЕСТОПОЛОЖЕНИЕ")
	 
	  self.result= 1
	  
	   
	   
	   
	   
     def task_generator(self):
	  yield Task ('post',url='https://qp.ru/realty/prodau_dachi',refresh_cache=True,network_try_count=100)
	  yield Task ('post',url='https://qp.ru/realty/sdau_dachi',refresh_cache=True,network_try_count=100)
     
   
     def task_post(self,grab,task):
	  for elem in grab.doc.select('//a[@target="_self"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	  yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	 
     def task_page(self,grab,task):
	  try:         
	       pg = grab.doc.select(u'//ul[@class="pagination js-next"]/li/a[contains(@title,"Следующая")]')
	       u = grab.make_url_absolute(pg.attr('href'))
	       yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	  except DataNotFound:
	       print('*'*50)
	       print '!!!','NO PAGE NEXT','!!!'
	       print('*'*50)
			 
    
   
     def task_item(self, grab, task):
	  
	  try:
	       klass =  grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[@class="controls"][1]').text().split(' / ')[0]
	  except IndexError:
	       klass = ''	  

	  try:
	       mesto = grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[contains(text(),"район")]').text().split(' / ')[1]
	  except IndexError:
	       mesto =''
	       
	  try:
	       try:
		    if  grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[contains(text(),"район")]').exists()==True:
			 punkt = grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[2]').text().split(' / ')[2]			 
		    else:
			 punkt = grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[2]').text().split(' / ')[1]
	       except IndexError:
		    punkt = grab.doc.select(u'//span[contains(text(),"Населённый пункт")]/following::div[2]').text()
	  except IndexError:
	       punkt = ''	       
	   
	  try:
	       ter =  grab.doc.select(u'//span[contains(text(),"Район города")]/following::div[2]').text()
	  except IndexError:
	       ter =''
	  try:
	       uliza = grab.doc.select(u'//span[contains(text(),"Улица")]/following::div[2]').text()
	  except IndexError:
	       uliza = ''
	  try:
	       dom = grab.doc.select(u'//span[contains(text(),"Номер дома")]/following::div[2]').text()
	  except IndexError:
	       dom = ''
	       
	  try:
	       tip = grab.doc.select(u'//span[contains(text(),"Объект")]/following::div[2]').text()
	  except IndexError:
	       tip = ''
	  try:
	       naz = grab.doc.select(u'//span[contains(text(),"Тип дома")]/following::div[2]').text()
	  except IndexError:
	       naz =''
	  
	  try:
	       price = grab.doc.select(u'//div[@class="btn-group price-dropdown js-dropdown-openhover"]/button').text()#.replace(' q',u' руб.')
	  except IndexError:
	       price =''
	  try: 
	       plosh = grab.doc.select(u'//span[contains(text(),"Общая площадь")]/following::div[2]').text()
	  except IndexError:
	       plosh=''
	  try:
	       ohrana = grab.doc.select(u'//span[contains(text(),"Этажность")]/following::div[2]').text()
	  except IndexError:
	       ohrana =''
	  try:
	       gaz =  grab.doc.select(u'//span[contains(text(),"Площадь участка")]/following::div[2]').text()
	  except IndexError:
	       gaz =''
	  try:
	       voda =  grab.doc.select(u'//h1').text()
	  except IndexError:
	       voda =''
	  try:
	       kanal = re.sub(u'^.*(?=санузел)','', grab.doc.select(u'//*[contains(text(), "санузел")]').text())[:7].replace(u'санузел',u'есть')
	  except IndexError:
	       kanal =''
	  try:
	       elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
	  except IndexError:
	       elek =''
	  try:
	       teplo = grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[@class="controls"][1]').text().replace(' / ',', ')
	  except IndexError:
	       teplo =''
	  #time.sleep(1)
	  try:
	       opis = grab.doc.select(u'//span[contains(text(),"Дополнительная информация")]/following::div[2]').text() 
	  except IndexError:
	       opis = ''
	  try:
	       try:
		    lico = grab.doc.select(u'//i[@class="fa fa-user"]/following::div[1]/div').text()
	       except IndexError:
		    lico = grab.doc.select(u'//div[@class="comment"]').text()
	  except IndexError:
	       lico = ''
	  try:
	       co = grab.doc.select(u'//i[@class="fa fa-user"]/following::div[1]/div').text()
	       if "едвижимост" in co:
		    comp = co
	       else:
		    comp=''
	  except IndexError:
	       comp = ''
	  try:
	       data1 = grab.doc.select(u'//dt[contains(text(),"Обновленно")]/following-sibling::dd[1]').text() 
	  except IndexError:   
	       data1 = ''
	  try: 
	       data = grab.doc.select(u'//i[@class="fa fa-calendar "]/following-sibling::text()').text()
	  except IndexError:
	       data=''
	  
	  url1 = re.sub('[^\d]','',task.url)
	  try:
	       phone_url = 'https://qp.ru/viewadvert/ShowPhones?id='+url1+'&datatype=json'    
	       headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
		         'Accept-Encoding': 'gzip,deflate',
		         'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
		         #'Cookie': 'QPSC4='+url1+'.'+url1,
		         'Host': 'qp.ru',
		         'Referer': task.url,
		         'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0',
		         'X-Requested-With': 'XMLHttpRequest'}
	       g2 = grab.clone(headers=headers,proxy_auto_change=True)
	       g2.request(headers=headers,url=phone_url) 
	       phone = ', '.join(g2.doc.json["phones"])
	       print 'Phone-OK'
	       del g2
	  except (IndexError,KeyError,ValueError,GrabConnectionError,GrabNetworkError,GrabTimeoutError):
	       del g2
	       phone = random.choice(list(open('../phone.txt').read().splitlines()))
	       
	  try:
	       if 'prodau' in task.url:
		    oper = u'Продажа' 
	       elif 'sdau' in task.url:
		    oper = u'Аренда'     
	  except IndexError:
	       oper = ''	       
     
	  
	       

   
	  projects = {'sub': klass,
                     'adress': mesto,
                      'terit':ter.replace(punkt,''), 
                      'punkt':punkt, 
                      'ulica':uliza,
                      'dom':dom,
                      'tip':tip,
                      'naz':naz,
                      'cena': price,
                      'plosh': plosh,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis': opis,
	              'oper':oper,
                      'url': task.url,
                      'phone': re.sub('[^\d\+\,]','',phone),
                      'lico':lico.replace(comp,''),
                      'company': comp,
                      'data':data,
                      'data1':data1}
     
     
	  yield Task('write',project=projects,grab=grab)
     
     
     
     def task_write(self,grab,task):
	  #time.sleep(1)
	  print('*'*100)	       
	  print  task.project['sub']
	  print  task.project['punkt']
	  print  task.project['adress']
	  print  task.project['terit']
	  print  task.project['ulica']
	  print  task.project['dom']
	  print  task.project['tip']
	  print  task.project['naz']
	  #print  task.project['klass']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanaliz']
	  print  task.project['electr']
	  print  task.project['ohrana']
	  print  task.project['opis']
	  print  task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	  print  task.project['teplo']
     
	  
	  

	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['adress'])
	  self.ws.write(self.result, 3, task.project['terit'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 10, task.project['tip'])
	  self.ws.write(self.result, 17, task.project['naz'])
	  #self.ws.write(self.result, 8, task.project['klass'])
	  self.ws.write(self.result, 12, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  self.ws.write(self.result, 16, task.project['ohrana'])
	  self.ws.write(self.result, 19, task.project['gaz'])
	  self.ws.write(self.result, 38, task.project['voda'])
	  self.ws.write(self.result, 23, task.project['kanaliz'])
	  self.ws.write(self.result, 24, task.project['electr'])
	  self.ws.write(self.result, 39, task.project['teplo'])
	  self.ws.write(self.result, 29, task.project['opis'])
	  self.ws.write(self.result, 30, u'КУПИ.РУ')
	  self.ws.write_string(self.result, 31, task.project['url'])
	  self.ws.write(self.result, 32, task.project['phone'])
	  self.ws.write(self.result, 33, task.project['lico'])
	  self.ws.write(self.result, 34, task.project['company'])
	  self.ws.write(self.result, 35, task.project['data'])
	  #self.ws.write(self.result, 32, task.project['data1'])
	  self.ws.write(self.result, 37, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 11, task.project['oper'])
	  print('*'*100)
	  #print self.sub
	  print 'Ready - '+str(self.result)
	  print 'Tasks - %s' % self.task_queue.size() 
	  print task.project['oper']
	  print('*'*100)
	  self.result+= 1
	  
	 
	  
	  
	  
	  #if self.result > 10:
	       #self.stop()
	       
	  #if str(self.result) == str(self.num):
	       #self.stop()		    


bot = QP_Com(thread_number=10, network_try_limit=1000)
#bot.setup_queue('mongo', database='qpZem',host='192.168.10.200')
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
try:
     bot.run()
except KeyboardInterrupt:
     pass
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')  
command = 'mount -a'
os.system('echo %s|sudo -S %s' % ('1122', command))
time.sleep(2)
workbook.close()
print('Done')

       
     
     
     
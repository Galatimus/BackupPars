#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
import time
import xlsxwriter
from datetime import datetime,timedelta
import random
import sys
import os
#sys.path.append('../')
#from datar import int_value_from_ru_month
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'zemm/0001-0072_00_У_003-0001_M-SAKH.xlsx')


class Farpost_Zem(Spider):
     def prepare(self):
	  #self.f = page
	  self.ws = workbook.add_worksheet(u'Sakh_Земля')
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
	  self.ws.write(0, 30, u"КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
	  self.ws.write(0, 31, u"ВИД_ПРАВА")
	  self.ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
	  self.result= 1
	  self.headers ={'Accept': 'application/json, text/javascript, */*; q=0.01'}	  
	
	       
    
     def task_generator(self):
	  for x in range(1,64):#59
               yield Task ('post',url='https://dom.sakh.com/land//list%d'%x+'/',refresh_cache=True,network_try_count=100)
	  
	  
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//div[@class="actions noprint"]/following-sibling::a'):
               ur = grab.make_url_absolute(elem.attr('href'))  
               #print ur
               yield Task('item', url=ur,refresh_cache=True,network_try_count=100)

     def task_item(self, grab, task):
	 
	  ray = ''          
	  try:
	       punkt= grab.doc.select(u'//div[@class="seller"]/preceding-sibling::h4[1]').text().split(', ')[0]
	  except IndexError:
	       punkt = ''
	       
	  try:
	       ter= grab.doc.select(u'//div[@class="seller"]/preceding-sibling::h4[1]').text().split(', ')[1]
	  except IndexError:
	       ter =''
	       
	  try:
	       
	       uliza = grab.doc.select(u'//div[@class="seller"]/preceding-sibling::h4[1]').text().split(', ')[2]
	       #else:
		    #uliza = ''
	  except IndexError:
	       uliza = ''
	       
          try:
               dom = grab.doc.select(u'//div[@id="breadcrumbs"]/div/a[contains(@href,"land")]/span').text().split(' ')[0]
          except IndexError:
               dom = ''
	       
	  try:
	       trassa = grab.doc.select(u'//label[contains(text(),"Шоссе:")]/following-sibling::p').text().split(', ')[0]
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//label[contains(text(),"Шоссе:")]/following-sibling::p').text().split(', ')[1]
	  except IndexError:
	       udal = ''
	       
	  try:
	       price = grab.doc.select(u'//div[@class="sum"][1]').text()#.replace(u'a',u'р.')
	  except IndexError:
	       price = ''

	  try:
	       plosh = grab.doc.select(u'//h4[1]').text().split(': ')[1].split(' в ')[0]
	  except IndexError:
	       plosh = ''

	  try:
	       vid = grab.doc.select(u'//h4[1]').text().split(', ')[1]
	  except DataNotFound:
	       vid = '' 
	       
	       
	  try:
	       ohrana = grab.doc.select(u'//div[@class="sum"][2]').text().replace('(','').replace(')','')
	  except DataNotFound:
	       ohrana =''
	  try:
	       gaz = grab.doc.select(u'//div[@class="seller"]/preceding-sibling::h4[1]').text()#.split(u'Обновлено: ')[1]
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
	       teplo = grab.doc.select(u'//div[@class="name"]/a').text()
	  except DataNotFound:
	       teplo =''
	  try:
	       opis = grab.doc.select(u'//div[@class="fulltext"]').text() 
	  except IndexError:
	       opis = ''

	  try:
	       lico = grab.doc.select(u'//div[@class="name"]/text()').text()
	  except IndexError:
	       lico = ''
	       
	  try:
               if 'sell' in task.url:
	            oper = u'Продажа' 
               elif 'lease' in task.url:
	            oper = u'Аренда'
               else:
	            oper=''
          except IndexError:
	       oper = ''
	       
	  try:
	       d = grab.doc.select(u'//div[@class="stat"]/div[1]/text()').text().split(',')[0].replace(u'Добавлено ','')
	       if 'вчера' in d:
		    data = '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1))
	       elif 'сегодня' in d:
		    data = (datetime.today().strftime('%d.%m.%Y'))	       
	       elif '2017' in d:
	            data = d+'2017'
	       elif '2016' in d:
	            data = d+'2016'
	       else:
	            data = d+'.2018'
	  except IndexError:
	       data = ''
	       
	       
	  url1 = re.sub('[^\d]','',task.url)
          phone_url = 'https://dom.sakh.com/dom/usrajax.php?action=get-phone&id='+url1+'&type=dom-offers'      
          headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
	            'Accept-Encoding': 'gzip, deflate, br',
	            'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
	            'Content-Length': '42',
	            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
	            'Cookie': 'PHPSESSID=h24vbsnvtp1dfnfetsdupefpd3',
	            'Host': 'dom.sakh.com',
	            'Referer': task.url,
	            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0', 
	            'X-Requested-With': 'XMLHttpRequest'}
	  g2 = grab.clone(headers=headers,proxy_auto_change=True)
     
	  #for ph in range(1,5):
	       #try:               
		    #time.sleep(1)
		    #g2.request(post=[('action','get-phone'), ('id', url1),('type', 'dom-offers')],headers=headers,url=phone_url)
		    #print g2.response.body
		    ##phone =  re.sub('[^\d\+]','',re.findall('em class=(.*?)/em>',g2.response.body)[0]) 
		    #phone =  re.sub('[^\d\+]','',g2.doc.rex_text(u'em class=(.*?)/em>'))
		    #print 'Phone-OK'
		    #del g2
		    #break  
	       #except (IndexError,GrabConnectionError,GrabNetworkError,GrabTimeoutError):
		    #g2.change_proxy()
		    #print 'Change proxy'+' : '+str(ph)+' / 5'
		    #g2 = grab.clone(headers=headers,timeout=2, connect_timeout=2,proxy_auto_change=True) 
	  #else:
	       #try:
		    #phone = grab.doc.select(u'//em[@class="text"]').text()
	       #except IndexError:
	  phone = random.choice(list(open('links/sphone.txt').read().splitlines()))	  		       
		    
	  
						   
	       
	  projects = {'url': task.url,
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
                      'lico':lico,
                      'company':oper,
	              'phone':phone,
                      'data':data[:10].replace('..201','.2018')}
	  
	  
	        
          
	  yield Task('write',project=projects,grab=grab)
            
     def task_write(self,grab,task):
	  print('*'*50)
	  #print  task.project['sub']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['teritor']
	  print  task.project['ulica']
	 
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
	  print  task.project['opis']
	  print task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  #print  task.project['company']
	  print  task.project['data']
	  print  task.project['teplo']

	  
	  #global result
	  self.ws.write(self.result, 0, u'Сахалинская область')
	  self.ws.write(self.result, 1, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 3, task.project['teritor'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  #self.ws.write(self.result, 9, task.project['dom'])
	  self.ws.write(self.result, 7, task.project['trassa'])
	  self.ws.write(self.result, 8, task.project['udal'])
	  #self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write_string(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 14, task.project['vid'])
	  self.ws.write(self.result, 32, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 27, task.project['teplo'])
	  self.ws.write(self.result, 11, task.project['ohrana'])
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'Market.sakh.com')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 9, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  print  task.project['company']
	  print('*'*50)	       
	  self.result+= 1
	       
	       
	       
	  #if self.result > 10:
	       #self.stop()

     
bot = Farpost_Zem(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
command = 'mount -a'
os.system('echo %s|sudo -S %s' % ('1122', command))
time.sleep(2)
workbook.close()
print('Done') 

time.sleep(5)
os.system("/home/oleg/pars/small/sakh_com.py")








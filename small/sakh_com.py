#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
import random
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


     
#proxy_lines= open('../tipa.txt').read().splitlines()

workbook = xlsxwriter.Workbook(u'comm/0001-0072_00_C_003-0001_M-SAKH.xlsx')





class Farpost_Com(Spider):
     def prepare(self):
	  #self.f = page
	  self.ws = workbook.add_worksheet(u'Sakh')
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
	  for x in range(1,60):#54
               yield Task ('post',url='https://dom.sakh.com/business//list%d'%x+'/',refresh_cache=True,network_try_count=100)
	  
			      
     def task_post(self,grab,task):    
	  for elem in grab.doc.select(u'//div[@class="actions noprint"]/following-sibling::a'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	  #yield Task("page", grab=grab,network_try_count=100)
     
     def task_page(self,grab,task):
	  try:
	       pg = grab.doc.select(u'//a[@class="step"][contains(text(),"следующая")]')
	       u = grab.make_url_absolute(pg.attr('href'))
	       yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	  except DataNotFound:
	       print('*'*100)
	       print '!!!','NO PAGE NEXT','!!!'
	       print('*'*100)
	       logger.debug('%s taskq size' % self.task_queue.size())      
        
     def task_item(self, grab, task):
	 
	  try:
	       ray = grab.doc.select(u'//div[@class="seller"]/preceding-sibling::h4[1]').text().split(', ')[0]
	  except IndexError:
	       ray = ''          
	
	       
	  try:
	       trassa = grab.doc.select(u'//h3').text()
		#print rayon
	  except IndexError:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//div[@class="area"]').text().split(u'Этаж: ')[1].split('/')[0]
	  except IndexError:
	       udal = ''
          try:
	       try:
                    seg =  grab.doc.select(u'//div[@class="area"]').text().split(u'Этаж: ')[1].split('/')[1]
	       except IndexError:
		    seg =  grab.doc.select(u'//div[@class="area"]').text().split(u'Этажей: ')[1]
          except IndexError:
               seg = ''	       
	       
	  try:
               price = grab.doc.select(u'//div[@class="sum"][1]').text()#.replace(u'a',u'р.')
	  except IndexError:
	       price = ''
	       
	  try:
	       plosh = grab.doc.select(u'//div[@class="area"]/text()[1]').text().replace(u'Площадь: ','')#.split(u'Площадь: ')[1].split(u'Этаж: ')[0].replace(u'Этажей: ','').replace(seg,'')
	  except IndexError:
	       plosh = '' 
	  try:
	       cena_za = grab.doc.select(u'//div[@class="name"]/text()').text()
	  except IndexError:
	       cena_za = '' 
	       
	  
	  try:
	       ohrana = grab.doc.select(u'//div[@class="name"]/a').text()
	  except IndexError:
	       ohrana =''
	  try:
	       gaz = grab.doc.select(u'//div[@class="seller"]/preceding-sibling::h4[1]').text()#.split(u'Обновлено: ')[1]
	  except IndexError:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//div[@class="stat"]/div[1]/span').attr('title')+'.2018'
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.select(u'//h2').text()+', '+grab.doc.select(u'//h3').text()
	  except IndexError:
	       kanal =''
	  try:
	       elek = grab.doc.rex_text(u'LatLng(.*?)zoom').split(', ')[0].replace('(','')
	  except IndexError:
	       elek =''
	  try:
	       teplo = grab.doc.rex_text(u'LatLng(.*?)zoom').split(', ')[1].replace(')','').replace(',','')
	  except IndexError:
	       teplo =''  
		    
	  try:
	       opis = grab.doc.select(u'//div[@class="fulltext"]').text() 
	  except IndexError:
	       opis = ''
	       
	 	       
	  
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
	       
	  #try:
	       #phone = grab.doc.select(u'//em[@class="text"]').text()
	  #except IndexError:
	  #phone = ''	       
	  
	  
	  
	  
	  #for p in range(1,51):
	       #proxy_line = proxy_lines[random.randint(1, len(proxy_lines)-1)].strip()
	       #session = requests.session()
	       #session.mount('http://', requests.adapters.HTTPAdapter(pool_connections = 1, pool_maxsize = 0, max_retries = 0))
	       #session.mount('https://', requests.adapters.HTTPAdapter(pool_connections = 1, pool_maxsize =0, max_retries = 0))
	       #session.proxies = {
                    #'http' : proxy_line,
                    #'https' : proxy_line
                    #}

		    
	       ##proxiesDict = {'http':  proxy_line,'https': proxy_line}
	       
	       #try:
		    #ad_id= re.sub(u'[^\d]','',task.url)
		    #headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
				   #'Accept-Encoding': 'gzip, deflate, br',
				   #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
				   #'Content-Length': '42',
				   #'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
				   #'Cookie': 'PHPSESSID=h24vbsnvtp1dfnfetsdupefpd3',
				   #'Host': 'dom.sakh.com',
				   #'Referer': task.url,
				   #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0', 
				   #'X-Requested-With': 'XMLHttpRequest'}       
		    #url_ph =  'https://dom.sakh.com/dom/usrajax.php?action=get-phone&id='+ad_id+'&type=dom-offers' 
		    #payload = (('action', 'get-phone'), ('id', ad_id),('type', 'dom-offers'))
		    
		    #r = session.post(url_ph,allow_redirects=True,headers=headers,data=payload,timeout=2)
		    
		    #print('Подключение удалось | ' + proxy_line)
		    #print r.text
		    #phone =  re.sub('[^\d]','',re.findall('em class=(.*?)/em>', r.text)[0])
		    #r.close()
		    #session.close()
		    ##Adapter.close()
		    #break
	       #except Exception as exception:
		    ##print(exception)
		    #print('Подключение не удалось. Переподключение. ' + proxy_line +' : '+str(p)+' / 50')
		    ##r.close()
		    #session.close()
		
	  #else:
	       #try:
                    #phone = grab.doc.select(u'//em[@class="text"]').text()
               #except IndexError:
                    #phone = ''	       
		    
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
     
	  #for ph in range(1,3):
	       #try:               
		    #time.sleep(1)
		    #g2.request(post=[('action','get-phone'), ('id', url1),('type', 'dom-offers')],headers=headers,url=phone_url)
		    #print g2.doc.body
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
	  phone = random.choice(list(open('links/sphone.txt').read().splitlines()))	  
	       
	
	       
	  projects = {'url': task.url,
                      'rayon': ray,
                      'trassa': trassa,
                      'udal': udal,
	              'segment': seg,
                      'cena': price,
                      'plosh':plosh,
	              'phone':phone[:12],
	              'cena_za': cena_za,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
	              'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'operazia':oper,
                      'data':data[:10].replace('..201','.2018')}
	  
          
	
	       
	  #ad_id= re.sub(u'[^\d]','',task.url)
	  #link = 'https://dom.sakh.com/dom/usrajax.php?action=get-phone&id='+ad_id+'&type=dom-offers'
	  #headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
                    #'Accept-Encoding': 'gzip, deflate, br',
                    #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
                    #'Content-Length': '42',
                    #'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    #'Cookie': 'PHPSESSID=h24vbsnvtp1dfnfetsdupefpd3',
                    #'Host': 'dom.sakh.com',
                    #'Referer': task.url,
                    #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0', 
                    #'X-Requested-With': 'XMLHttpRequest'}  
	  #it = Grab(timeout=50, connect_timeout=100)
	  #it.setup(url=link,post=[('action','get-phone'), ('id', ad_id),('type', 'dom-offers')],headers=headers)
	  #yield Task('phone',grab=it,project=projects,refresh_cache=True,network_try_count=100)
	  #del it
	  
	
	       
     #def task_phone(self, grab, task):
	  #try:
	       #print grab.response.body
	       #phone1 =  re.sub('[^\d\+]','',grab.doc.rex_text(u'em class=(.*?)/em>'))
	       #print phone1
	  #except IndexError:
	       #phone1 = ''     



	  #yield Task('write',project=task.project,grab=grab)
	  
	  yield Task('write',project=projects,grab=grab)
            
     def task_write(self,grab,task):
	  print('*'*50)
	  
	  print  task.project['rayon']
	  #print  task.phone
	  #print  task.project['phone']
	  #print  task.proj['ulica']
	  #print  task.proj['dom']
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
	  print  task.project['phone']
	  print  task.project['data']
          print  task.project['teplo']
	  
	  #global result
	  self.ws.write(self.result, 0, u'Сахалинская область')
	  self.ws.write(self.result, 2, task.project['rayon'])
	  self.ws.write(self.result, 35, task.project['teplo'])
	  self.ws.write(self.result, 34, task.project['electr'])
	  self.ws.write(self.result, 33, task.project['kanaliz'])
	  #self.ws.write(self.result, 21, task.phone)
	  self.ws.write(self.result, 21, task.project['phone'])
	  self.ws.write(self.result, 9, task.project['trassa'])
	  self.ws.write(self.result, 15, task.project['udal'])
	  self.ws.write(self.result, 16 , task.project['segment'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  self.ws.write(self.result, 23, task.project['ohrana'])
	  self.ws.write(self.result, 24, task.project['gaz'])
	  self.ws.write(self.result, 30, task.project['voda'])
	  #self.ws.write(self.result, 22, self.lico)
	  self.ws.write(self.result, 22, task.project['cena_za'])
	  self.ws.write(self.result, 19, u'Market.sakh.com')
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
	       
	       
	       
	  #if self.result > 5:
	       #self.stop()

     
bot = Farpost_Com(thread_number=5,network_try_limit=1000)
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
os.system("/home/oleg/pars/small/roszem.py")








#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
import math
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






i = 0
l= open('links/zem1.txt').read().splitlines()
dc = len(l)
page = l[i]
oper = u'Продажа'

while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class MK_Zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       while True:
		    try:
			 time.sleep(3)
			 g = Grab(timeout=10, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 self.sub = g.doc.select(u'//span[@class="current"]').text()
			 print self.sub
			 try:
			      self.num = re.sub('[^\d]','',g.doc.select(u'//span[@class="b-all-offers"]').text())
			      self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 except IndexError:
			      self.pag=0
			      self.num=0      
			 print self.sub,self.num,self.pag
			 break
		    except(GrabTimeoutError,GrabNetworkError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 #g.change_proxy()
			 continue
	       
	       self.workbook = xlsxwriter.Workbook(u'zem/Mirkvartir_%s' % bot.sub + u'_Земля_'+oper+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Mirkvartir_Земля')
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
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'?p=%d'%x,refresh_cache=True,network_try_count=100)
		 
		 
	  def task_post(self,grab,task):
	       if grab.doc.select(u'//h2[@class="nearby-header"]').exists()==True:
                    links = grab.doc.select(u'//h2[@class="nearby-header"]/preceding::div[@class="item"]/a[1]')
               else:
	            links = grab.doc.select(u'//div[@class="item"]/a[1]')
               for elem in links:
	            ur = grab.make_url_absolute(elem.attr('href'))  
	            #print ur
	            yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
              
	     
	     
	  def task_item(self, grab, task):
	       
	       try:
		    ray = grab.doc.select(u'//a[@class="js-popup-select popup-select Province-popup"]/following::span[1]').text()
		  #print ray 
	       except DataNotFound:
		    ray = ''          
	       try:
		    punkt= grab.doc.select(u'//a[@class="js-popup-select popup-select City-popup"]/following::span[1]').text()
	       except IndexError:
		    punkt = ''
		    
	       try:
		    ter= grab.doc.select(u'//a[@class="js-popup-select popup-select InhabitedPoint-popup"]/following::span[1]').text()
	       except IndexError:
		    ter =''
		    
	       try:
		    
		    uliza = grab.doc.select(u'//a[@class="js-popup-select popup-select Street-popup"]/following::span[1]').text()
		    #else:
			 #uliza = ''
	       except IndexError:
		    uliza = ''
		    
	       try:
		    dom = grab.doc.select(u'//label[contains(text(),"Адрес:")]/following-sibling::p/a[contains(@href,"houseId")]').text()
	       except DataNotFound:
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
		    price = grab.doc.select(u'//p[@class="price"]').text().replace(' ','')+u' р.'
	       except DataNotFound:
		    price = ''
		    
	       
		    
	       try:
		    plosh = grab.doc.select(u'//label[contains(text(),"Площадь:")]/following-sibling::p').text()
	       except DataNotFound:
		    plosh = ''
		    
	       
	       
	       
		    
	       try:
		    vid = grab.doc.select(u'//label[contains(text(),"Адрес:")]/following-sibling::p').text()
	       except DataNotFound:
		    vid = '' 
		    
		    
	       try:
		    ohrana =  grab.doc.select(u'//label[contains(text(),"Безопасность:")]/following-sibling::p').text().replace(u'охрана',u'есть')
	       except DataNotFound:
		    ohrana =''
	       try:
		    z =  grab.doc.select(u'//label[contains(text(),"Коммуникации:")]/following-sibling::p').text()
		    if z.find(u'газ')>=0:
			 gaz='есть'
		    else:
			 gaz=''
	       except DataNotFound:
		    gaz =''
	       try:
		    v =  grab.doc.select(u'//label[contains(text(),"Коммуникации:")]/following-sibling::p').text()
		    if v.find(u'вода')>=0:
			 voda='есть'
		    else:
			 voda=''
	       except DataNotFound:
		    voda =''
	       try:
		    k =  grab.doc.select(u'//label[contains(text(),"Коммуникации:")]/following-sibling::p').text()
		    if k.find(u'канализация')>=0:
			 kanal='есть'
		    else:
			 kanal=''
	       except DataNotFound:
		    kanal =''
	       try:
		    lk =  grab.doc.select(u'//label[contains(text(),"Коммуникации:")]/following-sibling::p').text()
		    if lk.find(u'электричество')>=0:
			 elek='есть'
		    else:
			 elek=''
	       except DataNotFound:
		    elek =''
	       try:
		    teplo = re.sub('[^\d\/]','',grab.doc.select(u'//p[@class="price"]/following-sibling::p[1]').text()).replace('/',u' руб.')
	       except DataNotFound:
		    teplo =''
		    
	         
		   
			 
	       try:
		    opis = grab.doc.select(u'//div[@class="clear"]/following-sibling::p').text() 
	       except DataNotFound:
		    opis = ''
		    

	       try:
		    lico = grab.doc.select(u'//span[@class="phones"]/following-sibling::text()').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//a[@rel="nofollow"]').text().replace(u'Показать телефон','')
	       except IndexError:
		    comp = ''
		    
	       try:
		    
		    conv = [ (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
			     (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1))),
			     (u'августа', '.08.2017'),(u'мая', '.05.2017'),(u'ноября', '.11.2017'),(u'января', '.01.2018'),(u'февраля', '.02.2018'),
			     (u'марта', '.07.2017'),(u'апреля', '.04.2017'),(u'октября', '.10.2017'),
			     (u'июля', '.07.2017'),(u'июня', '.06.2017'),(u'сентября', '.09.2017'),(u'декабря', '.12.2017')]
			     #(u'Июн', '.06.2015'),(u'июн', '.06.2015'),
			     #(u'Май', '.05.2015'),(u'май', '.05.2015'),
			     #(u'Апр', '.04.2015'),(u'апр', '.04.2015')]
		    dt= grab.doc.rex_text(u'Опубликовано: (.*?)в ').replace(' (','')
		    data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(' ','').replace(u'более3-хмесяце','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=92)))
		 #print data
	       except IndexError:
		    data = ''
		    
	       ad_id= re.sub(u'[^\d]','',task.url)
	       try:
                    ad_phone = grab.doc.select(u'//span[@class="phone"]/a').attr('key')
	       except IndexError:
		    ad_phone=''
               link = grab.make_url_absolute('/EstateOffers/AwesomeDecryptPhone/?offerId='+ad_id+'&encryptedPhone='+ad_phone)      
	       headers ={'Accept': '*/*',
	                 'Accept-Encoding': 'gzip,deflate',
	                 'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
	                 'Cookie': 'aliasIds='+ad_id+'.'+ad_phone,
	                 'Host': 'dom.mirkvartir.ru',
	                 'Referer': task.url,
	                 'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0', 
	                 'X-Requested-With' : 'XMLHttpRequest'}
	       g2 = grab.clone(headers=headers,proxy_auto_change=True)
	  
	       for ph in range(1,5):
		    try:               
			 g2.request(headers=headers,url=link)
			 #print g2.response.body
			 phone = g2.response.json["normalizedPhone"]
			 #phone =  re.sub('[^\d\+]','',re.findall('em class=(.*?)/em>',g2.response.body)[0]) 
			 #phone =  re.sub('[^\d\+]','',g2.doc.rex_text(u'em class=(.*?)/em>'))
			 print 'Phone-OK'
			 del g2
			 break  
		    except (IndexError,GrabConnectionError,GrabNetworkError,GrabTimeoutError,ValueError,TypeError,AttributeError):
			 g2.change_proxy()
			 print 'Change proxy'+' : '+str(ph)+' / 5'
			 g2 = grab.clone(headers=headers,timeout=2, connect_timeout=2,proxy_auto_change=True) 
	       else:
		    phone = ''		    
			 
	       
	       
							
		    
	       projects = {'url': task.url,
		           'sub': self.sub,
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
	                   'phone':phone,
		           'kanaliz': kanal,
		           'electr': elek,
		           'teplo': teplo,
		           'opis':opis,
		            'lico':lico,
		           'company':comp,
		           'data':data.replace('20172017г.','2017')}
	                   
	       
	       #try:
	
		    #ad_id= re.sub(u'[^\d]','',task.url)
		    #ad_phone = grab.doc.select(u'//span[@class="phone"]/a').attr('key')
		    #link = grab.make_url_absolute('/EstateOffers/AwesomeDecryptPhone/?offerId='+ad_id+'&encryptedPhone='+ad_phone)
		    #headers ={'Accept': '*/*',
			      #'Accept-Encoding': 'gzip,deflate',
			      #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			      #'Cookie': 'sessid='+ad_id+'.'+ad_phone,
			      #'Host': 'mirkvartir.ru',
			      #'Referer': task.url,
			      #'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0', 
			      #'X-Requested-With' : 'XMLHttpRequest'}
		    #gr = Grab()
		    #gr.setup(url=link,headers=headers)
		    #yield Task('phone',grab=gr,project=projects,refresh_cache=True,network_try_count=100)
	       #except IndexError:
	            #yield Task('phone',grab=grab,project=projects)	    

	  #def task_phone(self, grab, task):
	       #try:
		    #phone = grab.doc.rex_text(u'normalizedPhone":"(.*?)"')
	       #except (IndexError,ValueError,GrabNetworkError,GrabTimeoutError,IOError):
		    #phone = ''

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
	       print  task.project['vid']
	       print  task.project['data']
	      
	       
	       #global result
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 7, task.project['trassa'])
	       self.ws.write(self.result, 8, task.project['udal'])
	       self.ws.write(self.result, 9, oper)
	       self.ws.write(self.result, 10, task.project['cena'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 30, task.project['vid'])
	       self.ws.write(self.result, 15, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 18, task.project['electr'])
	       self.ws.write(self.result, 11, task.project['teplo'])
	       self.ws.write(self.result, 20, task.project['ohrana'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'MIRKVARTIR.RU')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 28, task.project['data'])
	       self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	       
	       print('*'*50)
	       print self.sub
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print '***',i+1,'/',dc,'***'
	       print oper
	       print('*'*50)
	       
	       self.result+= 1
		    
		    
		    
	       #if self.result > 50:
		    #self.stop()
     
	  
     bot = MK_Zem(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
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
     bot.workbook.close()
     print('Done')
   
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break








#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import random
import os
import math
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)






i = 0
l= open('Links/Zemm.txt').read().splitlines()
page = l[i]





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class QP_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	      
               for p in range(1,51):
                    try:
                         time.sleep(2)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http') 
                         g.go(self.f)
			 self.sub = g.doc.select(u'//a[@class="nav-link btn-region float-left jsc-region-selector"]/text()').text().replace('/','=')
                         self.num = re.sub('[^\d]','',g.doc.select(u'//strong[@class="items-count"]').text())
	                 self.pag = int(math.ceil(float(int(self.num))/float(30)))                         
                         print self.sub,self.pag,self.num
			 del g
                         break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.sub = ''
		    self.pag = 1
		    self.num = 1	       
		    
	       self.workbook = xlsxwriter.Workbook(u'zem/Qp_%s' % bot.sub + u'_Земля_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Qp_Земля')
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
	       for x in range(self.pag+1):
                    yield Task ('post',url=self.f+'?offset='+str(x*30),refresh_cache=True,network_try_count=50)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@target="_self"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=50)
	      
            
	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[contains(text(),"район")]').text().split(' / ')[1]
	       except IndexError:
	            mesto =''
		    
	       try:
		    try:
	                 punkt = grab.doc.select(u'//span[contains(text(),"Регион")]/following::div[2]').text().split(' / ')[2]
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
                    tip = grab.doc.select(u'//span[contains(text(),"Категория земель")]/following::div[2]').text()
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//span[contains(text(),"Использование")]/following::div[2]').text()
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//span[contains(text(),"Расстояние до города")]/following::div[2]').text()
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//div[@class="btn-group price-dropdown js-dropdown-openhover"]/button').text()#.replace(' q',u' руб.')
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//span[contains(text(),"Площадь")]/following::div[2]').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//span[contains(text(),"Этажность")]/following::div[2]').text()
               except IndexError:
                    ohrana =''
               try:
                    gaz =  grab.doc.select(u'//dt[@itemprop="name"][contains(text(),"Материал стен")]/following-sibling::dd').text()
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//h4').text().split('/')[1]
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
		    else:
			 oper = ''
	       except IndexError:
	            oper = ''		       
     
	
               projects = {'sub': self.sub,
	                  'adress': mesto,
	                   'terit':ter, 
	                   'punkt':punkt, 
	                   'ulica':uliza,
	                   'dom':dom,
	                   'tip':tip,
	                   'naz':naz,
	                   'klass': klass,
	                   'cena': price,
	                   'plosh': plosh,
	                   'ohrana':ohrana,
	                   'gaz': gaz,
	                   'voda': voda,
	                   'kanaliz': kanal,
	                   'operacia': oper,
	                   'electr': elek,
	                   'teplo': teplo,
	                   'opis': opis,
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
	       print  task.project['klass']
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
	       self.ws.write(self.result, 13, task.project['tip'])
	       self.ws.write(self.result, 14, task.project['naz'])
	       self.ws.write(self.result, 8, task.project['klass'])
	       self.ws.write(self.result, 10, task.project['cena'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 20, task.project['ohrana'])
	       self.ws.write(self.result, 15, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 18, task.project['electr'])
	       self.ws.write(self.result, 30, task.project['teplo'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       self.ws.write(self.result, 23, u'КУПИ.РУ')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 28, task.project['data'])
	       #self.ws.write(self.result, 32, task.project['data1'])
	       self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 9, task.project['operacia'])
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+self.num
	       print 'Tasks - %s' % self.task_queue.size()
	       print '***',i+1,'/',len(l),'***'
	       print task.project['operacia']
	       print('*'*100)
	       self.result+= 1
	       
	      
	       
	       
	       
	       #if self.result > 10:
	            #self.stop()
		    
	       if str(self.result) == str(self.num):
		    self.stop()		    


     bot = QP_Com(thread_number=10, network_try_limit=100)
     #bot.setup_queue('mongo', database='qpZem',host='192.168.10.200')
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     bot.run()
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
       
     
     
     
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

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






i = 0
l= open('links/zem.txt').read().splitlines()
page = l[i]
oper = u'Продажа'




while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class Ners_zem(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	      
	      
               while True:
                    try:
                         time.sleep(2)
			 g = Grab(timeout=10, connect_timeout=10)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http')			 
                         g.go(self.f)
                         self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="title_count"]/span').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(10)))
		         self.sub = g.doc.select(u'//div[@id="siteMenu"]/ul/li[1]/a').text()
			 print self.pag,self.num,self.sub
			 del g
                         break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
		    
	       self.workbook = xlsxwriter.Workbook(u'zem/Ners_%s' % bot.sub + u'_Земля_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Ners__Земля')
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
	       self.conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
		    (u' мая ',u'.05.'),(u' июня ',u'.06.'),
		    (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
		    (u' января ',u'.01.'),(u' декабря ',u'.12.'),
		    (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
		    (u' февраля ',u'.02.'),(u' октября ',u'.10.'),
		    (u'сегодня,', (datetime.today().strftime('%d.%m.%Y'))),
		    (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]	       
	      
	       self.result= 1
	       
                
                
                
                
	  def task_generator(self):
	       for x in range(self.pag+1):
                    yield Task ('post',url=self.f+'?start='+str(x*10),refresh_cache=True,network_try_count=10)
          
        
	  def task_post(self,grab,task):
	       #links = open('ners_zem1.txt', 'a')
	       for elem in grab.doc.select('//div[@class="media-body"]/div/following-sibling::a'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur    
		    #links.write("%s\n" % ur)
	       #links.close()		    
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=10)
	      
            
	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//dt[contains(text(),"Район области:")]/following-sibling::dd').text()
	       except IndexError:
	            mesto =''
		    
	       try:
	            if self.sub == u"Москва":
	                 punkt= u"Москва"
	            elif self.sub == u"Санкт-Петербург":
	                 punkt= u"Санкт-Петербург"
	            elif self.sub == u"Севастополь":
	                 punkt= u"Севастополь"
	            else:
	                 punkt = grab.doc.select(u'//dt[contains(text(),"Населенный пункт:")]/following-sibling::dd').text()
	       except IndexError:
	            punkt = ''	       
		
               try:
                    ter =  grab.doc.select(u'//dt[contains(text(),"Шоссе:")]/following-sibling::dd').text()
               except IndexError:
                    ter =''
               try:
		    uliza = grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd[contains(text(),"ул")]').text().split(', ')[0]
               except IndexError:
                    uliza = ''
		    
	       tip = ''

               try:
                    dom = re.sub('[^\d]','',grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text().split(', ')[1])
               except IndexError:
                    dom = ''
		    
               
               try:
                    naz = grab.doc.select(u'//span[contains(text(),"Использование")]/following::div[2]').text()
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//dt[contains(text(),"Расстояние от МКАД")]/following-sibling::dd').text()
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//div[@class="price_value"]/text()').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//dt[contains(text(),"Площадь:")]/following-sibling::dd').text()
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
                    voda =  grab.doc.select(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//span[@class="price_for_unit"]').text()
               except IndexError:
                    kanal =''
               try:
                    elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
               except DataNotFound:
                    elek =''
               try:
                    teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
               except DataNotFound:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//div[@class="info mb-4"]').text() 
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//a[@class="profile_link"]').text()
	       except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//a[@class="firm_link"]').text()
               except IndexError:
                    comp = ''
               try:
                    d = grab.doc.select(u'//div[contains(text(),"Дата размещения:")]').text().replace(u'Дата размещения: ','') 
                    data1 = reduce(lambda d, r: d.replace(r[0], r[1]), self.conv, d)
               except IndexError:   
                    data1 = ''
	       try: 
	            dt = grab.doc.select(u'//div[contains(text(),"Дата обновления:")]').text().replace(u'Дата обновления: ','')#[:9]
                    data = reduce(lambda dt, r: dt.replace(r[0], r[1]), self.conv, dt)[:10]
	       except IndexError:
		    data=''
		    
	       
	       
	       	
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
	                   'electr': elek,
	                   'teplo': teplo,
	                   'phone': random.choice(list(open('../phone.txt').read().splitlines())),
	                   'opis': opis,
	                   'url': task.url,
	                   'lico':lico,
	                   'company': comp,
	                   'data':data,
	                   'data1':data1}
	       #try:
		    #key = re.sub('[^\d]','',grab.doc.select(u'//div[@class="notes_id"]/b').text()) 
		    #user = re.sub('[^\d]','',grab.doc.select(u'//button[@id="get_phone"]/@data-u').text())
		    #pkey = re.sub('[^\d]','',grab.doc.select(u'//body/@data-stat-id').text())
		    #link = task.url.split(u'object')[0]#+u'ru'
		    #url_ph = 'https://ru.ners.ru/ajax/?module=notes_get_phone&notes_id='+key+'&user_id='+user+'&db_importer_id=0'+'&stat_id='+pkey+'&is_my_notes=0'+'&app_alias=ru&json=1'
		    #headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
			      #'Accept-Encoding': 'gzip,deflate',
			      #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			      #'Cookie': 'sessid='+key+'.'+pkey,
			      #'Host': 'ru.ners.ru',
		              #'Origin': link,
			      #'Referer': task.url,
			      #'X-Requested-With' : 'XMLHttpRequest',
			      #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:23.0) Gecko/20131011 Firefox/23.0'}		    
		    #gr = Grab()
		    #gr.setup(url=url_ph,headers=headers)	            
		    #yield Task('phone',grab=gr,project=projects,refresh_cache=True,network_try_count=10)
               #except IndexError:
	            #yield Task('phone',grab=grab,project=projects)	       
	       
	       
	  #def task_phone(self, grab, task):
	       #try:
		    ##phone =  re.sub('[^\d]','',re.findall("tel:(.*?)>",grab.response.body)[0])
		    #phone = grab.response.json["phone"]
		    #print grab.response.body
	       #except IndexError:
		    #phone = ''
	         
	       #yield Task('write',project=task.project,phone=phone,grab=grab)
	       yield Task('write',project=projects,grab=grab)
	  
	  
	  
	  def task_write(self,grab,task):
	       if task.project['cena'] <> '':
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
		    print  task.project['teplo']
		    print  task.project['ohrana']
		    print  task.project['opis']
		    print  task.project['url']
		    print  task.project['phone']
		    print  task.project['lico']
		    print  task.project['company']
		    print  task.project['data']
		    print  task.project['data1']
		   
	       
		    
		    
	  
		    self.ws.write(self.result, 0, task.project['sub'])
		    self.ws.write(self.result, 1, task.project['adress'])
		    self.ws.write(self.result, 7, task.project['terit'])
		    self.ws.write(self.result, 2, task.project['punkt'])
		    self.ws.write(self.result, 4, task.project['ulica'])
		    self.ws.write(self.result, 5, task.project['dom'])
		    self.ws.write(self.result, 3, task.project['tip'])
		    self.ws.write(self.result, 14, task.project['naz'])
		    self.ws.write(self.result, 8, task.project['klass'])
		    self.ws.write(self.result, 10, task.project['cena'])
		    self.ws.write(self.result, 12, task.project['plosh'])
		    self.ws.write(self.result, 20, task.project['ohrana'])
		    self.ws.write(self.result, 15, task.project['gaz'])
		    self.ws.write(self.result, 31, task.project['voda'])
		    self.ws.write(self.result, 11, task.project['kanaliz'])
		    self.ws.write(self.result, 18, task.project['electr'])
		    #self.ws.write(self.result, 24, task.project['teplo'])
		    self.ws.write(self.result, 22, task.project['opis'])
		    self.ws.write(self.result, 23, u'Национальная единая риэлторская сеть')
		    self.ws.write_string(self.result, 24, task.project['url'])
		    self.ws.write(self.result, 25, task.project['phone'])
		    self.ws.write(self.result, 26, task.project['lico'])
		    self.ws.write(self.result, 27, task.project['company'])
		    self.ws.write(self.result, 28, task.project['data1'])
		    self.ws.write(self.result, 29, task.project['data'])
		    self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
		    self.ws.write(self.result, 9, oper)
		    print('*'*100)
		    
		    print 'Ready - '+str(self.result)+'/'+self.num
		    logger.debug('Tasks - %s' % self.task_queue.size()) 
		    print '***',i+1,'/',len(l),'***'
		    print oper
		    print('*'*100)
		    self.result+= 1
		    
		   
		    
		    
		    
		    #if self.result > 10:
			 #self.stop()	       
	

	 

     bot = Ners_zem(thread_number=5, network_try_limit=100)
     bot.load_proxylist('../ivan.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     #command = 'mount -a'
     #os.system('echo %s|sudo -S %s' % ('1122', command))
     #time.sleep(2)
     bot.workbook.close()
     print('Done')
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
       
     
time.sleep(5)
os.system("/home/oleg/pars/ners/comm.py")
     
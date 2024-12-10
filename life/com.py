#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import os

#from PIL import Image
#import pytesseract
import math
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



#g.proxylist.load_file(path='/home/oleg/Proxy/tipa.txt',proxy_type='http')



i = 0
l= open('links/com_prod.txt').read().splitlines()
page = l[i]
oper = u'Продажа'




while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class move_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	      
               while True:
                    try:
                         time.sleep(2)
			 g = Grab(timeout=20, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                         g.go(self.f)
			 
			 self.sub = g.doc.select(u'//span[@id="showRegionBlock"]').text()
			 try:
                              self.num = re.sub('[^\d]','',g.doc.select(u'//a[@id="nextPage"]/preceding-sibling::a[1]').text())
	                 except IndexError:
			      self.num = 0
                         print self.sub,self.num
			 del g
                         break
			
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
		    
	       self.workbook = xlsxwriter.Workbook(u'com/Life-Realty_%s' % bot.sub + u'_Коммерческая_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'life-realty_Коммерческая')
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
	      
	       self.result= 1
	       
                
                
                
                
	  def task_generator(self):
	       for x in range(1,int(self.num)+1):
	            yield Task ('post',url=self.f+'?page=%d'%x,refresh_cache=True,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//td[@class="txt"]/a'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)

	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//div[@class="value"]/a[contains(text(),"район")]').text()
	       except IndexError:
	            mesto =''
		    
	       try:
	            punkt = grab.doc.rex_text(u'Населенный пункт: (.*?)<br>')
	       except IndexError:
	            punkt = ''	       
		
               try:
                    ter= grab.doc.rex_text(u'Район: (.*?)</span>').replace('<span class="grey">','')
		    
               except IndexError:
                    ter =''
	       try:
		   
		    uliza= grab.doc.rex_text(u'Адрес: (.*?)<br>').split(', ')[0]
		    #t2=0
	            #for w1 in r1.split(','):
			 #t2+=1
			 #for x in range(len(ul)):
			      #if ul[x] in w1:
				   #uliza = re.sub('\d+$', '',r1.split(',')[t2-1]).replace(' д','')
				   #break
		    #print uliza
	       except (IndexError,UnboundLocalError):
		    uliza =''
               try:
                    dom = grab.doc.rex_text(u'Адрес: (.*?)<br>').split(', ')[1]
		    #dom = re.compile(r'[0-9]+$',re.S).search(dm).group(0)
               except (IndexError,AttributeError):
                    dom = ''
	         
               try:
                    tip = grab.doc.select(u'//h4[@class="header7"]').text().split(' - ')[1]
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//h1').text().replace('Продается ','').replace('Сдается ','')
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//span[contains(text(),"Этаж")]/following::div[2]').text()
               except IndexError:
                    klass = ''
               try:
		    price = grab.doc.select(u'//div[@class="card_price"]').text()
		    #except IndexError:
			 #price = grab.doc.select(u'//span[@id="price-total-0"]').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.rex_text(u'Площадь помещения: (.*?)</sup>').replace('<sup>','')
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
               except DataNotFound:
                    elek =''
               try:
                    teplo = grab.doc.rex_text(u'var address = "(.+?)";')
               except DataNotFound:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//h4[contains(text(),"Дополнительная информация")]/following-sibling::text()').text() 
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//div[@class="c_face"]').text()
	       except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//span[contains(text(),"Агентство недвижимости")]/preceding::a[1]').text()
               except IndexError:
                    comp = ''
               
	       try: 
		    conv = [(u' августа',u'.08.'), (u' июля',u'.07.'),
			    (u' мая',u'.05.'),(u' июня',u'.06.'),
			    (u' марта',u'.03.'),(u' апреля',u'.04.'),
			    (u' января',u'.01.'),(u' декабря',u'.12.'),
			    (u' сентября',u'.09.'),(u' ноября',u'.11.'),
			    (u' февраля',u'.02.'),(u' октября',u'.10.'),
		            (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
		            (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]
	            d = grab.doc.select(u'//div[@class="card_date"]').text().split(u' добавлено ')[1]
		    data = reduce(lambda d, r: d.replace(r[0], r[1]), conv, d)[:11]
		    
	       except IndexError:
		    data=''
		    
	       
	       try:
                    phone = re.sub('[^\d\,]','',grab.doc.select(u'//div[@class="c_phone"]').text())
               except IndexError:
	            phone = ''
          
	       
		    
     
	
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
	                   'opis': opis,
	                   'url': task.url,
	                   'phone': phone,
	                   'lico':lico,
	                   'company': comp,
	                   'data':re.sub('[^\d\.]','',data)}
	                   
	  
	  
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
	       self.ws.write(self.result, 8, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       self.ws.write(self.result, 13, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 14, task.project['ohrana'])
	       self.ws.write(self.result, 16, task.project['gaz'])
	       self.ws.write(self.result, 35, task.project['voda'])
	       self.ws.write(self.result, 22, task.project['kanaliz'])
	       self.ws.write(self.result, 23, task.project['electr'])
	       self.ws.write(self.result, 34, task.project['teplo'])
	       self.ws.write(self.result, 25, task.project['opis'])
	       self.ws.write(self.result, 26, u'Life-Realty.ru')
	       self.ws.write_string(self.result, 27, task.project['url'])
	       self.ws.write(self.result, 28, task.project['phone'])
	       self.ws.write(self.result, 29, task.project['lico'])
	       self.ws.write(self.result, 30, task.project['company'])
	       self.ws.write(self.result, 31, task.project['data'])
	       #self.ws.write(self.result, 32, task.project['data1'])
	       self.ws.write(self.result, 32, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 33, oper)
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*100)
	       self.result+= 1
	       
	      
	       
	       
	       
	       #if self.result >= 50:
	            #self.stop()	       
	


     bot = move_Com(thread_number=5, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
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
          if oper == u'Продажа':
               i = 0
               l= open('links/com_arenda.txt').read().splitlines()
               dc = len(l)
               page = l[i]
               oper = u'Аренда'
          else:
               break
       
     
     
     
#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import math
import os
import re
import random
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)






i = 0
l= open('Links/Com_Prod.txt').read().splitlines()
page = l[i]
oper = u'Продажа'




while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class Nndv_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	      
               for p in range(1,50):
                    try:
                         time.sleep(1)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')			 
                         g.go(self.f)
			 if g.doc.select(u'//strong[contains(text(),"не найдено")]').exists()==False:
                              self.num = g.doc.select(u'//section[@id="content"]/h1/strong').number()
	                      self.pag = int(math.ceil(float(int(self.num))/float(50)))
                              self.sub = g.doc.select(u'//span[contains(text(),"Быстрый переход:")]/following-sibling::a[2]').text().replace('/','+')
                              print self.sub,self.pag,self.num
			      del g
                              break
			 else:
			      self.sub=''
			      self.pag=1
			      self.num=1
			      del g
			      break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
               else:
		    self.sub=''
		    self.pag=1
		    self.num=1		    
		    
	       self.workbook = xlsxwriter.Workbook(u'com/Nndv_%s' % bot.sub + u'_Коммерческая_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Nndv_Коммерческая')
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
	       
	       self.result= 1
	       
                
                
                
                
	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=page+'%d'%x,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//h3/a[contains(@href, "html")]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	      
            
	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//label[contains(text(),"Район")]/following-sibling::div/a').text()
	       except IndexError:
	            mesto =''
		    
	       try:
		    try:
	                 punkt = grab.doc.select(u'//label[contains(text(),"Город")]/following-sibling::div/a').text()
		    except IndexError:
			 punkt = grab.doc.select(u'//label[contains(text(),"Населенный пункт")]/following-sibling::div/a').text()
	       except IndexError:
	            punkt = ''	       
		
               try:
                    ter =  grab.doc.select(u'//title').text()
               except IndexError:
                    ter =''
               try:
                    uliza = grab.doc.select(u'//label[contains(text(),"Адрес")]/following-sibling::div/a').text()
               except IndexError:
                    uliza = ''
               try:
                    dom = grab.doc.select(u'//meta[@itemprop="latitude"]').attr('content')
               except IndexError:
                    dom = ''
		    
               try:
                    tip = grab.doc.select(u'//meta[@itemprop="longitude"]').attr('content')
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//label[contains(text(),"Объект:")]/following-sibling::div').text().split(' ')[0]
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Класс ")]/preceding-sibling::div').text()
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//label[contains(text(),"Стоимость:")]/following-sibling::div').text().replace(u'rur ','')+u' р.'
               except IndexError:
                    price =''
               try: 
                    plosh = re.sub('[^\d\m]','',grab.doc.select(u'//label[contains(text(),"Объект:")]/following-sibling::div').text())
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//div[@class="offer-detail__section-item section_type_floor"]/div[1]').text()
               except IndexError:
                    ohrana =''
               try:
                    gaz =  grab.doc.select(u'//div[@class="offer-detail__section-item section_type_building-floors"]/div[1]').text()
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Год постройки")]/preceding-sibling::div').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Высота потолков")]/preceding-sibling::div').text()
               except IndexError:
                    kanal =''
               try:
                    elek = grab.doc.select(u'//div[@class="offer-detail__section-item-header"][contains(text(),"Состояние ")]/preceding-sibling::div').text()
               except IndexError:
                    elek =''
               try:
                    teplo = self.sub+','+punkt+','+grab.doc.select(u'//label[contains(text(),"Адрес")]/following-sibling::div/a').text()+' '+mesto
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//label[contains(text(),"Комментарий:")]/following-sibling::div').text() 
	       except IndexError:
	            opis = ''
               
               try:
                    comp = grab.doc.select(u'//meta[@itemprop="legalName"]').attr('content')
               except IndexError:
                    comp = ''
	       try:
	            lico = grab.doc.select(u'//meta[@itemprop="name"]').attr('content')
	       except IndexError:
	            lico = ''   
               
	       try: 
	            data = grab.doc.select(u'//label[contains(text(),"Размещено:")]/following-sibling::div').text()#.split(' ')[0], '%Y.%d.%m')
	       except IndexError:
		    data=''
	       
	       try:
		    phone = random.choice(list(open('../phone.txt').read().splitlines()))
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
	                   'data':data}#.strftime('%d.%m.%Y')}
	                   
	  
	  
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
	       self.ws.write(self.result, 33, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 34, task.project['dom'])
	       self.ws.write(self.result, 35, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       #self.ws.write(self.result, 10, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       #self.ws.write(self.result, 13, task.project['ohrana'])
	       #self.ws.write(self.result, 14, task.project['gaz'])
	       #self.ws.write(self.result, 15, task.project['voda'])
	       #self.ws.write(self.result, 17, task.project['kanaliz'])
	       #self.ws.write(self.result, 18, task.project['electr'])
	       self.ws.write(self.result, 24, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'Необходимая недвижимость')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 28, oper)
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       print 'Tasks - %s' % self.task_queue.size() 
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*100)
	       self.result+= 1
	       
	      
	       
	       
	       
	       #if self.result > 10:
	            #self.stop()	       

	 

     bot = Nndv_Com(thread_number=5, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')    
     #command = 'mount -a'
     #os.system('echo %s|sudo -S %s' % ('1122', command))
     time.sleep(2)
     bot.workbook.close()
     print('Done')    
     i=i+1
     try:
          page = l[i]
     except IndexError:
          if oper == u'Продажа':
               i = 0
               l= open('Links/Com_Arenda.txt').read().splitlines()
               dc = len(l)
               page = l[i]
               oper = u'Аренда'
          else:
               break
       
     
     
     
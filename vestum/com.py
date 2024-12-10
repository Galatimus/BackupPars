#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
from sub import conv
import os
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
                         time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                         g.go(self.f)
			 dt = g.doc.select(u'//span[@class="arrow"]/span').text()
			 self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(u' крайский ',' ')
			 self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="catalog-counter"]').text())
			 self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 print self.sub,self.pag,self.num
			 del g
			 break
			 
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
		    
	       self.workbook = xlsxwriter.Workbook(u'com/Vestum_%s' % bot.sub + u'_Коммерческая_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'vestum_Коммерческая')
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
                    yield Task ('post',url=self.f+'?page=%d'% x,refresh_cache=True,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[contains(@class,"card")]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	      
            
	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//div[@id="content-address"]/a[contains(text(),"район")]').text()
	       except IndexError:
	            mesto =''
	       try:
		    mesto1 = grab.doc.select(u'//div[@id="content-address"]').text()
	       except IndexError:
	            mesto1 =''		    
	         
               try:
                    tip = grab.doc.select(u'//td[@class="param param-ext"][contains(text(),"Готовый ")]').text()#.split(' - ')[1]
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//div[@id="breadcrumbs"]/div/span[3]/a/span').text().replace(u'Купить ','').replace(u'Снять ','').replace(u'гостиницу',u'гостиница')
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//td[contains(text(),"Тип здания")]/following-sibling::td').text()
               except IndexError:
                    klass = ''
               try:
		    try:
                         price = grab.doc.select(u'//div[@class="price-main"]').text()
		    except IndexError:
			 price = grab.doc.select(u'//span[@class="price-main"]').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//td[contains(text(),"лощадь")]/following-sibling::td').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//td[contains(text(),"Этаж / этажей")]/following-sibling::td').text().split(' / ')[0].replace(u'/','-')
               except IndexError:
                    ohrana =''
               try:
                    gaz =  grab.doc.select(u'//td[contains(text(),"Этаж / этажей")]/following-sibling::td').text().split(' / ')[1].replace('/','-')
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//td[contains(text(),"Год постройки")]/following-sibling::td').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//h1').text()
               except IndexError:
                    kanal =''
               try:
                    elek = grab.doc.select(u'//div[@id="map"]').attr('data-x')
               except DataNotFound:
                    elek =''
               try:
                    teplo = grab.doc.select(u'//div[@id="map"]').attr('data-y')
               except DataNotFound:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//div[@class="text-info"]').text() 
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//a[@class="seller-name"]').text()
	       except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//div[@class="seller-parent an"]').text()
               except IndexError:
                    comp = ''
               

	       try:    
	            data = grab.doc.select(u'//td[contains(text(),"Актуально")]/following-sibling::td/span').attr('data-datetime')[:10].replace(' ','.')
	       except IndexError:
		    data =''
	       try:
		    data1 = grab.doc.select(u'//td[contains(text(),"Размещено")]/following-sibling::td/span').attr('data-datetime')[:10].replace(' ','.')
	       except IndexError:
		    data1=''
	       
	       try:
                    phone = re.sub('[^\d\,]','',grab.doc.select(u'//div[@id="seller-phone"]/span/div').attr('data-phone'))
               except IndexError:
	            phone = ''
          
	       
		    
     
	
               projects = {'sub': self.sub,
	                  'adress': mesto,
	                  'adress1': mesto1,
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
	                   'data':data,
	                   'data1':data1}
	       try:
		    ad= grab.doc.select(u'//div[@id="content-address"]').text()
		    link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ad
		    yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    yield Task('adres',grab=grab,project=projects)
		    
	  def task_adres(self, grab, task):
	       try:   
		    punkt= grab.doc.rex_text(u'LocalityName":"(.*?)"')
	       except IndexError:
		    punkt = ''
	       try:
		    ter=  grab.doc.rex_text(u'DependentLocalityName":"(.*?)"')
	       except IndexError:
		    ter =''
	       try:
		    uliza=grab.doc.rex_text(u'ThoroughfareName":"(.*?)"')
	       except IndexError:
		    uliza = ''
	       try:
		    dom = grab.doc.rex_text(u'PremiseNumber":"(.*?)"')
	       except IndexError:
		    dom = ''
	  
	       project2 ={'punkt':punkt,
	                  'teritor': ter,
	                  'ulica':uliza,
	                  'dom':dom}	  	  
	  
	       yield Task('write',project=task.project,proj=project2,grab=grab)
	  
	  
	  
	  def task_write(self,grab,task):
	       #time.sleep(1)
	       print('*'*100)	       
	       print  task.project['sub']
	       print  task.proj['punkt']
	       print  task.project['adress']
	       print  task.project['adress1']
	       print  task.proj['teritor']
	       print  task.proj['ulica']
	       print  task.proj['dom']
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
	       self.ws.write(self.result, 3, task.proj['teritor'])
	       self.ws.write(self.result, 2, task.proj['punkt'])
	       self.ws.write(self.result, 4, task.proj['ulica'])
	       self.ws.write(self.result, 5, task.proj['dom'])
	       #self.ws.write(self.result, 7, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       #self.ws.write(self.result, 16, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 15, task.project['ohrana'])
	       self.ws.write(self.result, 16, task.project['gaz'])
	       self.ws.write(self.result, 17, task.project['voda'])
	       self.ws.write(self.result, 33, task.project['kanaliz'])
	       self.ws.write(self.result, 34, task.project['electr'])
	       self.ws.write(self.result, 35, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'Вестум.RU')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 30, task.project['data'])
	       self.ws.write(self.result, 29, task.project['data1'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 28, oper)
	       self.ws.write(self.result, 24, task.project['adress1'])
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+self.num
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*100)
	       self.result+= 1
	       
	      
	       
	       
	       
	       #if self.result >= 10:
	            #self.stop()	       

	 

     bot = move_Com(thread_number=3, network_try_limit=1000)
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
       
     
     
     
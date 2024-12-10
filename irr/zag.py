#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import base64
from grab import Grab
import re
import time
import os
from datetime import datetime
import xlsxwriter
import math
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



#g.proxylist.load_file(path='/home/oleg/Proxy/tipa.txt',proxy_type='http')

i = 0
l= open('Links/Zagg.txt').read().splitlines()
dc = len(l)
page = l[i]


while True:
     print '********************************************',i+1,'/',dc,'*******************************************'	       
     class IRR_Zag(Spider):
	  def prepare(self):
	       #self.count = 1 
               self.f = page
               #self.link =l[i]
               #while True:
	            #try:
	                 #time.sleep(1)
			 #g = Grab(timeout=20, connect_timeout=20)
			 #g.proxylist.load_file(path='../ivan.txt',proxy_type='http')
			 #g.go(self.f)
			 #city = [ (u'кой',u'кая'),(u'области',u'область'),(u'ком',u'кий'),
			          #(u'Москве',u'Москва'),(u'Петербурге',u'Петербург'),
			          #(u'крае',u'край'),(u'республике ','')]
		    
			 #dt = g.doc.select(u'//span[@itemprop="name"]').text().replace('Все объявления в ','').replace('Все объявления во ','').replace('/','-')
		    
			 #self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), city, dt)
			 #try:
                              #self.num= re.sub('[^\d]', '',g.doc.select(u'//div[@class="listingStats"]').text().split('из ')[1])
                         #except IndexError:
	                      #self.num = '1'
			 #self.pag = int(math.ceil(float(int(self.num))/float(30)))
			 #print self.sub,self.num,self.pag
			 #del g
			 #break
		    #except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
			 #print g.config['proxy'],'Change proxy'
			 #g.change_proxy()
			 #del g
			 #continue
		    
               self.workbook = xlsxwriter.Workbook(u'zagg/IRR_Загород_'+str(i+1)+'.xlsx')
               self.ws = self.workbook.add_worksheet()
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
	       self.ws.write(0, 38, "КАТЕГОРИЯ_ЗЕМЛИ")
	       self.ws.write(0, 39, "МЕСТОПОЛОЖЕНИЕ")
	       self.result= 1
	       
            
            
            
              
    
	  def task_generator(self):
	       #for x in range(1,self.pag+1):
                    #link = self.f+'page'+str(x)+'/'
                    #yield Task ('post',url=link.replace(u'page1/',''),refresh_cache=True,network_try_count=100)
	       yield Task ('post',url= self.f,refresh_cache=True,network_try_count=100)
		    
	  def task_post(self,grab,task):
	  
	       if grab.doc.select(u'//h4[contains(text(),"Предложения рядом")]').exists()==True:
		    links = grab.doc.select(u'//h4[contains(text(),"Предложения рядом")]/preceding::a[contains(@class,"listing")]')
	       else:
		    links = grab.doc.select(u'//a[@class="listing__itemTitle"]')
	       for elem in links:
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	       yield Task('page', grab=grab,refresh_cache=True,network_try_count=100)
		    
	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//li[contains(@class,"active")]/following-sibling::li[1]/a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print 'no_page' 	  	  
        
	  def task_item(self, grab, task):
	       try:
		    sub =  grab.doc.rex_text('"address_region":"(.*?)"').decode("unicode_escape").replace(u'russia\/moskva-region\/',u'Москва').replace(u'russia\/tatarstan-resp\/',u'Татарстан')	     
	       except (IndexError,TypeError,ValueError):
		    sub = ''
	       except KeyError:
	            sub = u'Санкт-Петербург'	       
	       try:
                    ga = grab.doc.select(u'//h1').text()
                    t=0
                    for w in ga.split(','):
	                 t+=1
	                 if w.find(u'р-н')>=0:
	                      rayon = ga.split(', ')[t-1]
	                      break
                    if w.find(u'р-н')<0:
	                 rayon =''
               except IndexError:
                    rayon =''
	       try:
		    punkt = grab.doc.select(u'//i[@class="icon icon_spot"]/following-sibling::div').text().split(', ')[0].replace(u'Объявление на сайте продавца','')
		  #print punkt
	       except IndexError:
		    punkt = ''
		    
	       try:
		    ter =  grab.doc.select(u'//div[contains(text(),"Район города:")]/following-sibling::div[@class="propertyValue"]/a').text()
	       except IndexError:
		    ter ='' 
	       try:
	            uliza = grab.doc.select(u'//i[@class="icon icon_spot"]/following-sibling::div').text().split(', ')[1]
	         #print rayon
	       except IndexError:
		    uliza = ''
		    
               try:
	            dom = grab.doc.select(u'//i[@class="icon icon_spot"]/following-sibling::div').text().split(', ')[2]
	        #print rayon
	       except IndexError:
	            dom = ''
		   
	       try:
		    orentir = grab.doc.select(u'//div[contains(text(),"Направление:")]/following-sibling::div[@class="propertyValue"]').text()
		   #print rayon
	       except IndexError:
		    orentir = ''
               try:
                    tr = grab.doc.select(u'//h1').text()
                    t=0
                    for w in tr.split(','):
	                 t+=1
	                 if w.find(u'шоссе')>=0:
	                      trassa = tr.split(', ')[t-1].replace(u' шоссе','')
	                      break
                    if w.find(u'шоссе')<0:
	                 trassa =''
               except IndexError:
	            trassa = '' 
		   
	       try:
		    ud = grab.doc.select(u'//h1').text()
		    t=0
		    for w in ud.split(','):
			 t+=1
			 if w.find(u'км.')>=0:
			      udall = ud.split(', ')[t-1]
			      break
		    if w.find(u'км.')<0:
			 udall =''
	       except IndexError:
		    udall = ''
               try:
                    tip_ob = re.sub('[\d]','',grab.doc.select(u'//h1').text().split(u'кв.м')[0])
                #print rayon
               except IndexError:
                    tip_ob = ''
   
               try:
                    price = grab.doc.select('//div[@itemprop="price"]').text()
                 #print price + u' руб'	    
               except IndexError:
                    price = ''
   
               try:
		    plosh_str =re.sub('[^\d]','',grab.doc.select(u'//h1').text().split(u'кв.м')[0])+ u' м2'
	       except IndexError:
                    plosh_str = ''
     
               try:
                    komnat = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Количество комнат:")]').text().split(': ')[1]
                 #print price + u' руб'	    
               except IndexError:
                    komnat = ''
		   
               try:
                    etash = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Количество этажей:")]').text().split(': ')[1]
                 #print price + u' руб'	    
               except IndexError:
                    etash = ''
		    
               try:
                    mat = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Материал стен:")]').text().split(': ')[1]
                 #print rayon
               except IndexError:
                    mat = '' 
		   
               try:
                    god = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Год постройки/сдачи:")]').text().split(': ')[1]
                 #print rayon
               except IndexError:
                    god = ''
		   
               try:
		    uz = grab.doc.select(u'//h1').text()
		    u=0
		    for w in uz.split(','):
			 u+=1
			 if w.find(u'площадь участка')>=0:
			      plosh_uch = uz.split(', ')[u-1].replace(u'площадь участка ','')
			      break
		    if w.find(u'площадь участка')<0:
			 plosh_uch =''
	       except IndexError:
		    plosh_uch =''
		   
               try:
		    postroyki = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Гараж")]').text()
                 #print rayon
               except IndexError:
                    postroyki = ''
		   
               try:
		    gaz = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Газ в доме")]').text().replace(u'Газ в доме',u'есть')
                 #gazz = gaz.replace('True',u'есть')
               except IndexError:
                    gaz =''
		   
               try:
                    voda = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Водопровод")]').text().replace(u'Водопровод',u'есть')
                 #gazz = gaz.replace('True',u'есть')
               except IndexError:
                    voda =''
		   
               try:
                    kanal = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Канализация")]').text().replace(u'Канализация',u'есть')
                 #gazz = gaz.replace('True',u'есть')
               except IndexError:
                    kanal =''
		   
               try:
                    elekt = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Электричество")]').text().replace(u'Электричество (подведено)',u'есть')
               except IndexError:
                    elekt =''
   
   
               try:
                    teplo = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Отапливаемый")]').text().replace(u'Отапливаемый',u'есть')
                  #gazz = gaz.replace('True',u'есть')
               except IndexError:
                    teplo =''
		   
               try:
                    les = grab.doc.select(u'//i[@class="icon icon_spot"]/following-sibling::div').text()
                 #gazz = gaz.replace('True',u'есть')
               except IndexError:
                    les =''
		   
               try:
                    con = [(u' августа',u'.08.2019'), (u' июля',u'.07.2019'),
                         (u' мая',u'.05.2019'),(u' июня',u'.06.2019'),
                         (u' марта',u'.03.2019'),(u' апреля',u'.04.2019'),
                         (u' января',u'.01.2019'),(u' декабря',u'.12.2018'),
                         (u' сентября',u'.09.2019'),(u' ноября',u'.11.2018'),
                         (u' февраля',u'.02.2019'),(u' октября',u'.10.2018'),
                         (u'сегодня', (datetime.today().strftime('%d.%m.%Y')))]
	            d1 = grab.doc.select(u'//div[@class="productPage__createDate"]').text()
	            vodoem = reduce(lambda d1, r: d1.replace(r[0], r[1]), con, d1).replace(u'Размещено ','')
               except IndexError:
                    vodoem =''
		   
               try:
                    ohrana = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Охрана")]').text().replace(u'Охрана',u'есть')
                 #gazz = gaz.replace('True',u'есть')
               except IndexError:
                    ohrana =''
		   
               try:
                    opis = grab.doc.select(u'//h4[contains(text(),"Описание")]/following-sibling::p').text() 
               except IndexError:
                    opis = ''
		   
		   
               try:
                    phone = re.sub('[^\d]', '',base64.b64decode(grab.doc.select('//input[@name="phoneBase64"]').attr('value')))
               except (AttributeError,DataNotFound):
                    phone = ''
		   
               try:
                    lico = grab.doc.select(u'//div[@class="productPage__infoTextBold productPage__infoTextBold_inline"]/a[contains(@href,"user")]').text()
               except DataNotFound:
                    lico = ''
		   
               try:
                    comp = grab.doc.select(u'//div[@class="productPage__infoTextBold productPage__infoTextBold_inline"]/a[contains(@href,"russia")]').text()
                 #print rayon
               except DataNotFound:
                    comp = ''
		   
               try:
		    data = grab.doc.rex_text(u'date_create":"(.*?)"}').split(' ')[0].replace('-','.')
	       except DataNotFound:
		    data = ''
		   
               
		   
               try:
                    catzem = grab.doc.select(u'//li[@class="productPage__infoColumnBlockText"][contains(text(),"Категория земли:")]').text().split(': ')[1]
                 #print rayon
               except DataNotFound:
                    catzem = ''
		    
		    
	       if 'rent' in task.url:
		    oper = u'Аренда'
	       else:
	            oper = u'Продажа'	       
	  
	       projects = {'sub': sub,
		           'rayon': rayon,
		           'punkt': punkt,
	                   'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
	                   'orentir': orentir,
	                   'trassa': trassa,
	                   'udall': udall,
		           'object': tip_ob,
		           'cena': price,
		           'plostr': plosh_str,
		           'komnati': komnat,
		           'etach': etash,
		           'material': mat,
		           'god_postr': god,
		           'plouh': plosh_uch,
		           'postroyki': postroyki,
		           'gaz': gaz,
	                   'voda':voda,
		           'kanal': kanal,
		           'svet':elekt,
		           'teplo':teplo,
		           'les': les,
		           'bezop':ohrana,
		           'opis':opis,
	                   'url': task.url,
		           'phone':phone,
		           'lico':lico,
		           'company':comp,
	                   'operacia':oper,
	                   'data':data,
	                   'vodoem':vodoem,
	                   'zemlya':catzem}
	
	
	
	       yield Task('write',project=projects,grab=grab)
	
	 
	
	
	
	
          def task_write(self,grab,task):
	       
               print('*'*150)	       
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['orentir']
	       print  task.project['trassa']
	       print  task.project['udall']
	       print  task.project['object']
	       print  task.project['cena']
	       print  task.project['plostr']
	       print  task.project['komnati']
	       print  task.project['etach']
	       print  task.project['material']
	       print  task.project['god_postr']
	       print  task.project['plouh']
	       print  task.project['postroyki']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanal']
	       print  task.project['svet']
	       print  task.project['teplo']
	       print  task.project['les']
	       print  task.project['bezop']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['vodoem']
               print  task.project['zemlya']
	 
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 6, task.project['orentir'])
	       self.ws.write(self.result, 7, task.project['trassa'])
	       self.ws.write(self.result, 8, task.project['udall'])
	       self.ws.write(self.result, 10, task.project['object'])
	       self.ws.write(self.result, 11, task.project['operacia'])
	       self.ws.write(self.result, 12, task.project['cena'])	
	       self.ws.write(self.result, 14, task.project['plostr'])
	       self.ws.write(self.result, 15, task.project['komnati'])
	       self.ws.write(self.result, 16, task.project['etach'])
	       self.ws.write(self.result, 17, task.project['material'])
	       self.ws.write(self.result, 18, task.project['god_postr'])
	       self.ws.write(self.result, 19, task.project['plouh'])
	       self.ws.write(self.result, 20, task.project['postroyki'])
	       self.ws.write(self.result, 21, task.project['gaz'])
	       self.ws.write(self.result, 22, task.project['voda'])
	       self.ws.write(self.result, 23, task.project['kanal'])
	       self.ws.write(self.result, 24, task.project['svet'])
	       self.ws.write(self.result, 25, task.project['teplo'])
	       self.ws.write(self.result, 39, task.project['les'])
	       self.ws.write(self.result, 36, task.project['vodoem'])
	       self.ws.write(self.result, 28, task.project['bezop'])
	       self.ws.write(self.result, 29, task.project['opis'])
	       self.ws.write(self.result, 30, 'Из рук в руки')
	       self.ws.write_string(self.result, 31, task.project['url'])
	       self.ws.write(self.result, 32, task.project['phone'])
	       self.ws.write(self.result, 33, task.project['lico'])
	       self.ws.write(self.result, 34, task.project['company'])
	       self.ws.write(self.result, 35, task.project['data'])
	       self.ws.write(self.result, 37, datetime.today().strftime('%d.%m.%Y'))
               self.ws.write(self.result, 38, task.project['zemlya'])
	       print('*'*100)
	       
	       print('*'*50)
	       print 'Ready - '+str(self.result)
	       print 'Tasks - %s' % self.task_queue.size()
	       print '*',i+1,'/',len(l),'*'
	       print  task.project['operacia']
	       print('*'*50)
	       self.result+= 1
	       
	       #if self.result > 20:
		    #self.stop()
               #if str(self.result) == str(self.num):
		    #self.stop()		    
	       

	 
     
     bot = IRR_Zag(thread_number=5,network_try_limit=1000)
     #bot.setup_queue('mongo', database='IrrZag',host='192.168.10.200')
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     bot.run()    
     print('Wait 2 sec...')
     time.sleep(1)
     #print('Save it...')
     #p = os.system('echo %s|sudo -S %s' % ('1122', 'mount -a'))
     #print p     
     #time.sleep(2)     
     bot.workbook.close()
     print('Done!')     
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break
     
     
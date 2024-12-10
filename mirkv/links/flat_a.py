#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import math
import re
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
import os
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

i = 62
l= open('links/Kv_Arenda.txt').read().decode('cp1251').splitlines()
dc = len(l)
page = l[i] 
oper = u'Аренда'
     
#g = Grab(timeout=20, connect_timeout=20)

#g.proxylist.load_file(path='../tipa.txt',proxy_type='http')

while True:
     print '********************************************',i+1,'/',dc,'*******************************************'
     class MK_Kv(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	       while True:
		    try:
			 time.sleep(5)
			 g = Grab(timeout=20, connect_timeout=20)
			 g.go(self.f)
			 self.sub = g.doc.select(u'//span[@class="current"]').text()
			 try:
			      self.num = re.sub('[^\d]','',g.doc.select(u'//span[@class="b-all-offers"]').text())
			      self.pag = int(math.ceil(float(int(self.num))/float(20)))
			 except IndexError:
			      self.pag=0
			      self.num=0
		    
			 print self.sub,self.num,self.pag
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
		         print g.config['proxy'],'Change proxy'
		         g.change_proxy()
		         del g
		         continue
	       
	       self.workbook = xlsxwriter.Workbook(u'Kv/Mirkvartir_%s' % self.sub + u'_Жилье_'+oper+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Mirkvartir_Жилье')
	       self.ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	       self.ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	       self.ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
	       self.ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	       self.ws.write(0, 4, u"УЛИЦА")
	       self.ws.write(0, 5, u"ДОМ")
	       self.ws.write(0, 6, u"ОРИЕНТИР")
	       self.ws.write(0, 7, u"СТАНЦИЯ_МЕТРО")
	       self.ws.write(0, 8, u"ДО_МЕТРО_МИНУТ")
	       self.ws.write(0, 9, u"ПЕШКОМ_ТРАНСПОРТОМ")
	       self.ws.write(0, 10, u"ТИП_ОБЪЕКТА")
	       self.ws.write(0, 11, u"ОПЕРАЦИЯ")
	       self.ws.write(0, 12, u"СТОИМОСТЬ")
	       self.ws.write(0, 13, u"ЦЕНА_М2")
	       self.ws.write(0, 14, u"КОЛИЧЕСТВО_КОМНАТ")
	       self.ws.write(0, 15, u"ПЛОЩАДЬ_ОБЩАЯ")
	       self.ws.write(0, 16, u"ПЛОЩАДЬ_ЖИЛАЯ")
	       self.ws.write(0, 17, u"ПЛОЩАДЬ_КУХНИ")
	       self.ws.write(0, 18, u"ПЛОЩАДЬ_КОМНАТ")
	       self.ws.write(0, 19, u"ЭТАЖ")
	       self.ws.write(0, 20, u"ЭТАЖНОСТЬ")
	       self.ws.write(0, 21, u"МАТЕРИАЛ_СТЕН")
	       self.ws.write(0, 22, u"ГОД_ПОСТРОЙКИ")
	       self.ws.write(0, 23, u"РАСПОЛОЖЕНИЕ_КОМНАТ")
	       self.ws.write(0, 24, u"БАЛКОН")
	       self.ws.write(0, 25, u"ЛОДЖИЯ")
	       self.ws.write(0, 26, u"САНУЗЕЛ")
	       self.ws.write(0, 27, u"ОКНА")
	       self.ws.write(0, 28, u"СОСТОЯНИЕ")
	       self.ws.write(0, 29, u"ВЫСОТА_ПОТОЛКОВ")
	       self.ws.write(0, 30, u"ЛИФТ")
	       self.ws.write(0, 31, u"РЫНОК")
	       self.ws.write(0, 32, u"КОНСЬЕРЖ")
	       self.ws.write(0, 33, u"ОПИСАНИЕ")
	       self.ws.write(0, 34, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 35, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 36, u"ТЕЛЕФОН")
	       self.ws.write(0, 37, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 38, u"КОМПАНИЯ")
	       self.ws.write(0, 39, u"ДАТА_РАЗМЕЩЕНИЯ_ОБЪЯВЛЕНИЯ")
	       self.ws.write(0, 40, u"ДАТА_ПАРСИНГА")
	       #self.ws.write(0, 41, u"ДОП._ИНФОРМАЦИЯ")
	       
	       self.result= 1
            
            
            
              
    
	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'?p=%d'%x,network_try_count=100)
        
        
            
	  def task_post(self,grab,task):
	       if grab.doc.select(u'//h2[@class="nearby-header"]').exists()==True:
		    links = grab.doc.select(u'//h2[@class="nearby-header"]/preceding::div[@class="item"]/a')
	       else:
		    links = grab.doc.select(u'//div[@class="item"]/a')
	  
	       for elem in links:
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item',url=ur,refresh_cache=True,network_try_count=100)
	      
	    
        
        
     
	  def task_item(self, grab, task):
	       
	       
	       try:
		    ray = grab.doc.select(u'//a[@class="js-popup-select popup-select Province-popup"]/following::span[1]').text()
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
	       except IndexError:
		    uliza = ''
	       try:
		    dom = grab.doc.select(u'//label[contains(text(),"Адрес:")]/following-sibling::p/a[contains(@href,"houseId")]').text()
	       except DataNotFound:
		    dom = ''
		     
	       try:
		    orentir = grab.doc.select(u'//label[contains(text(),"Жилой комплекс:")]/following-sibling::p').text()
	       except DataNotFound:
		    orentir = ''              
		 
	       try:
		    metro = grab.doc.select(u'//label[contains(text(),"Метро:")]/following-sibling::p').text().split(', ')[0]
		 #print rayon
	       except IndexError:
		    metro = ''
		   
	       try:
		    metro_min = grab.doc.select(u'//span[@class="object_item_metro_comment"]').number()
		 #print rayon
	       except DataNotFound:
		    metro_min = ''
		   
	       try:
		    metro_tr = grab.doc.select(u'//span[@class="object_item_metro_comment"]').text().split(u'мин. ')[1]
	       except IndexError:
		    metro_tr = ''
		    
	       try:
		    if grab.doc.select(u'//label[contains(text(),"Квартира:")]').exists()==True:
			 tip_ob = u'Комната'
		    else:
			 tip_ob = u'Квартира' 
	       except DataNotFound:
		    tip_ob = ''
		    
	       
		   
	       try:
		    price = grab.doc.select(u'//p[@class="price"]/strong').text()+u' р.'
		 #print price + u' руб'	    
	       except IndexError:
		    price = ''
		   
	       try:
		    price_m = grab.doc.select(u'//small[contains(text(),"м²")]/ancestor::p').text()#.split(u'.')[0]
	       except IndexError:
		    price_m = ''
		     
	       try:
		    kol_komnat = grab.doc.select(u'//label[contains(text(),"Комнаты:")]/following-sibling::p').number()
		#print rayon
	       except DataNotFound:
		    kol_komnat = ''
     
	       
     
	       try:
		    plosh_ob = grab.doc.rex_text(u'<p>(.*?)sup2; ').replace('&','2')
		  #print rayon
	       except DataNotFound:
		    plosh_ob = ''
     
	       try:
		    plosh_gil = grab.doc.rex_text(u'жилая (.*?)sup2;').replace('&','2')
		  #print rayon
	       except DataNotFound:
		    plosh_gil = ''
		     
	       try:
		    plosh_kuh = grab.doc.rex_text(u'кухня (.*?)sup2;').replace('&','2')
		  #print rayon
	       except DataNotFound:
		    plosh_kuh = ''
		  
	       try:
		    plosh_com = grab.doc.select(u'//label[contains(text(),"Комнаты:")]/following-sibling::p/br/following-sibling::text()').text()
	       except DataNotFound:
		    plosh_com = ''
		    
	       try:
		    et = grab.doc.select(u'//label[contains(text(),"Этаж:")]/following-sibling::p').text().split(u' из ')[0]
		 #print price + u' руб'	    
	       except IndexError:
		    et = '' 
		   
	       try:
		    etagn = grab.doc.select(u'//label[contains(text(),"Этаж:")]/following-sibling::p').text().split(u' из ')[1]
		 #print price + u' руб'	    
	       except IndexError:
		    etagn = ''
		     
	       try:
		    mat = grab.doc.select(u'//label[contains(text(),"Дом:")]/following-sibling::p').text().split(', ')[0]
		 #print rayon
	       except IndexError:
		    mat = '' 
		   
	       try:
		    god = grab.doc.select(u'//label[contains(text(),"Дом:")]/following-sibling::p').text().split(', ')[1]
	       except IndexError:
		    god = ''
		     
	       try:
		    balkon = grab.doc.select(u'//th[contains(text(),"Балкон:")]/following-sibling::td[contains(text(),"балк")]').text()#.replace(u'нет','')
		 #print rayon
	       except DataNotFound:
		    balkon = ''
		   
	       try:
		    lodg = grab.doc.select(u'//th[contains(text(),"Балкон:")]/following-sibling::td[contains(text(),"лодж")]').text()
		 #print rayon
	       except DataNotFound:
		    lodg = ''
		   
	       try:
		    sanuzel = grab.doc.select(u'//th[contains(text(),"Санузел:")]/following-sibling::td').text().replace(u'нет','')
	       except DataNotFound:
		    sanuzel = ''
		     
		     
	       try:
		    okna = grab.doc.select(u'//th[contains(text(),"Вид из окна:")]/following-sibling::td').text()
	       except DataNotFound:
		    okna = ''
		   
	       #try:
		 #potolki = grab.doc.select(u'//div[contains(text(),"Высота потолков:")]/following-sibling::div[@class="propertyValue"]').text()
	       #except DataNotFound:
		   #potolki = ''
		   
	       try:
		    lift = grab.doc.select(u'//th[contains(text(),"Лифт:")]/following-sibling::td').text().replace(u'нет','')
	       except DataNotFound:
		    lift = ''
		  
	       try:
		    rinok = grab.doc.select(u'//th[contains(text(),"Тип дома:")]/following-sibling::td').text().split(', ')[0]
	       except DataNotFound:
		    rinok = ''
		   
	       try:
		    kons = re.sub(u'^.*(?=консьерж)','', grab.doc.select(u'//*[contains(text(), "консьерж")]').text())[:8].replace(u'консьерж',u'есть')
	       except DataNotFound:
		    kons = ''
		     
	       try:
		    opis = grab.doc.select(u'//div[@class="clear"]/following-sibling::p').text() 
	       except DataNotFound:
		    opis = ''
 
	       try:
		    lico = grab.doc.select(u'//h3[contains(text(),"Позвоните продавцу")]/following-sibling::p/text()').text()
	       except IndexError:
		    lico = ''
		    
	       try:
		    comp = grab.doc.select(u'//a[@rel="nofollow"]').text().replace(u'Показать телефон','')
		 #print rayon
	       except DataNotFound:
		    comp = ''
		    
	       try:
		    
		    conv = [ (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
		         (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1))),
		         (u'августа', '.08.2017'),(u'мая', '.05.2017'),(u'ноября', '.11.2016'),
		         (u'марта', '.03.2017'),(u'сентября', '.09.2017'),(u'октября', '.10.2017'),(u'января', '.01.2017'),(u'февраля', '.02.2017'),(u'апреля', '.04.2017'),
		         (u'июля', '.07.2017'),(u'июня', '.06.2017'),(u'декабря', '.12.2016')]
		    dt= grab.doc.rex_text(u'Опубликовано: (.*?)в ').replace(' (','')
		    data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(' ','').replace(u'более3-хмесяце','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=92)))
		 #print data
	       except IndexError:
		    data = ''
		    

	       projects = {'sub': self.sub,
		           'rayon': ray,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		           'dom': dom,
		           'metro': metro,
		           'udall': metro_min,
		           'tran': metro_tr,
		           'object': tip_ob,
		           'cena': price,
		           'cena_m': price_m,
		           'col_komnat': kol_komnat,
		           'plosh_ob':plosh_ob,
		           'plosh_gil': plosh_gil,
		           'plosh_kuh': plosh_kuh,
		           'plosh_com': plosh_com,
		           'etach': et,
		           'etashost': etagn,
		           'material': mat,
		           'god_postr': god,
		           'balkon': balkon,
		           'logia': lodg,
		           'uzel':sanuzel,
		           'okna': okna,
		           'lift':lift,
		           'rinok': rinok,
		           'kons':kons,
		           'opis':opis,
		           'url':task.url,
		           'lico':lico,
		           'company':comp,
		           'data':data}
	     
	       try:
		    #ad_id= re.sub(u'[^\d]','',task.url[-9:])
		    ad_id= re.sub(u'[^\d]','',task.url)
		    ad_phone = grab.doc.select(u'//span[@class="phone"]/a').attr('key')
		    link = grab.make_url_absolute('/EstateOffers/DecryptPhone?offerId='+ad_id+'&encryptedPhone='+ad_phone)
		    headers ={'Accept': '*/*',
			      'Accept-Encoding': 'gzip,deflate',
			      'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			      'Cookie': 'sessid='+ad_id+'.'+ad_phone,
			      'Host': 'mirkvartir.ru',
			      'Referer': task.url,
			      'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0', 
			      'X-Requested-With' : 'XMLHttpRequest'}
		    gr = Grab()
		    gr.setup(url=link)
		    yield Task('phone',grab=gr,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
	            yield Task('phone',grab=grab,project=projects)	     
	     
	     
	  def task_phone(self, grab, task):
	       try:
		    phone = grab.doc.rex_text(u'normalizedPhone":"(.*?)"')
	       except (IndexError,ValueError,GrabNetworkError,GrabTimeoutError,IOError):
		    phone = ''	  
	     
	     
	       yield Task('write',project=task.project,phone=phone,grab=grab)
	     
	     
	     
	     
	     
	     
	  def task_write(self,grab,task):
	       
	       print('*'*50)	       
	       print  task.project['sub']
	       print  task.project['rayon']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['metro']
	       print  task.project['udall']
	       print  task.project['tran']
	       print  task.project['object']
	       print  task.project['cena']
	       print  task.project['cena_m']
	       print  task.project['col_komnat']
	       print  task.project['plosh_ob']
	       print  task.project['plosh_gil']
	       print  task.project['plosh_kuh']
	       print  task.project['plosh_com']
	       print  task.project['etach']
	       print  task.project['etashost']
	       print  task.project['material']
	       print  task.project['god_postr']
	       print  task.project['balkon']
	       print  task.project['logia']
	       print  task.project['uzel']
	       print  task.project['okna']
	       print  task.project['lift']
	       print  task.project['rinok']
	       print  task.project['kons']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.phone
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       #print  task.project['tip_prod']
	 
	       self.ws.write(self.result, 0,task.project['sub'])
	       self.ws.write(self.result, 1,task.project['rayon'])
	       self.ws.write(self.result, 2,task.project['punkt'])
	       self.ws.write(self.result, 3,task.project['teritor'])
	       self.ws.write(self.result, 4,task.project['ulica'])
	       self.ws.write(self.result, 5,task.project['dom'])
	       self.ws.write(self.result, 7,task.project['metro'])
	       self.ws.write(self.result, 8,task.project['udall'])
	       self.ws.write(self.result, 9,task.project['tran'])
	       self.ws.write(self.result, 10,task.project['object'])
	       self.ws.write(self.result, 11,oper)
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 13, task.project['cena_m'])
	       self.ws.write(self.result, 14, task.project['col_komnat'])
	       self.ws.write(self.result, 15, task.project['plosh_ob'])
	       self.ws.write(self.result, 16, task.project['plosh_gil'])
	       self.ws.write(self.result, 17, task.project['plosh_kuh'])
	       self.ws.write(self.result, 18, task.project['plosh_com'])
	       self.ws.write(self.result, 19, task.project['etach'])
	       self.ws.write(self.result, 20, task.project['etashost'])
	       self.ws.write(self.result, 21, task.project['material'])
	       self.ws.write(self.result, 22, task.project['god_postr'])
	       self.ws.write(self.result, 24, task.project['balkon'])
	       self.ws.write(self.result, 25, task.project['logia'])
	       self.ws.write(self.result, 26, task.project['uzel'])
	       self.ws.write(self.result, 27, task.project['okna'])
	       self.ws.write(self.result, 30, task.project['lift'])
	       self.ws.write(self.result, 31, task.project['rinok'])
	       self.ws.write(self.result, 32, task.project['kons'])
	       self.ws.write(self.result, 33, task.project['opis'])
	       self.ws.write(self.result, 34, u'MIRKVARTIR.RU')
	       self.ws.write_string(self.result, 35, task.project['url'])
	       self.ws.write(self.result, 36, task.phone)
	       self.ws.write(self.result, 37, task.project['lico'])
	       self.ws.write(self.result, 38, task.project['company'])
	       self.ws.write(self.result, 39, task.project['data'])
	       self.ws.write(self.result, 40, datetime.today().strftime('%d.%m.%Y'))
	       #self.ws.write(self.result, 41, task.project['tip_prod'])
	       
	       print('*'*50)
	       print self.sub
	      
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',dc,'***'
	       print oper
	       print('*'*50)
	       self.result+= 1
	       
	   
	       #if self.result > 10:
		    #self.stop()
               if str(self.result) == str(self.num):
	            self.stop()		    
	
	
	       
     bot = MK_Kv(thread_number=3,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=10)
     bot.run()
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     try:
	  command = 'mount -t cifs //192.168.1.6/d /home/oleg/pars -o _netdev,sec=ntlm,auto,username=oleg,password=1122,file_mode=0777,dir_mode=0777'
	  os.system('echo %s|sudo -S %s' % ('1122', command))
	  time.sleep(2)
	  bot.workbook.close()
	  print('Done')
     except IOError:
	  time.sleep(10)
	  os.system('echo %s|sudo -S %s' % ('1122', 'mount -a'))
	  time.sleep(5)
	  bot.workbook.close()
	  print('Done!')     
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break

     
     
#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import logging
import os
#from cStringIO import StringIO
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






i = 0
l= open('links/com_prod.txt').read().splitlines()
page = l[i]
oper = u'Продажа'




while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'	       
     class Realtyvision_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]
	      
               while True:
                    try:
                         time.sleep(2)
			 g = Grab(timeout=20, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')			 
                         g.go(self.f)
			 for elem in g.doc.select('//p[@class="page_navigator"]/a'):
			      self.pag = elem.number()                         
	                 self.sub = g.doc.select(u'//div[contains(text(),"Недвижимость")]/a/text()').text().replace(u' области',u' область').replace(u'ой ',u'ая ').replace(u'ой ',u'ая ').replace(u'ом ',u'ий ')
                         print self.sub,self.pag
			 del g
                         break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
		    except AttributeError:
			 self.pag = 0
		    
	       self.workbook = xlsxwriter.Workbook(u'com/Realtyvision_%s' % bot.sub + u'_Коммерческая_'+oper+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Realtyvision_Коммерческая')
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
	       self.ws.write(0, 32, u"ДАТА_ОБНОВЛЕНИЯ")
	       self.ws.write(0, 33, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 34, u"ОПЕРАЦИЯ")
	       self.ws.write(0, 35, u"МЕСТОПОЛОЖЕНИЕ")
	       self.ws.write(0, 36, u"ЦЕНА_ЗА_М2")
	      
	       self.result= 1
	       
                
                
                
                
	  def task_generator(self):
	       for x in range(1,self.pag+1):
                    yield Task ('post',url=self.f+'?start='+str(x*50),refresh_cache=True,network_try_count=100)
          
        
	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@class="list_show_contacts"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)
	      
            
	 
        
	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//p[contains(text(),"Средняя цена на похожие объекты -")]/b').text().split(u', ')[3]
	       except IndexError:
	            mesto =''
		    
	       try:
	            punkt = grab.doc.select(u'//p[contains(text(),"Средняя цена на похожие объекты -")]/b').text().split(', ')[2]
	       except IndexError:
	            punkt = ''	       
		
               try:
                    t =  grab.doc.select(u'//div[@id="price_table"]/following::p[1]').text()
		    if u'Добавлено:' in t:
			 ter =  grab.doc.select(u'//div[@id="price_table"]/following::p[2]').text()
		    else:
			 ter = t
               except IndexError:
                    ter =''
               try:
                    uliza = grab.doc.select(u'//td[contains(text(),"Год постройки")]/following-sibling::td').text()
               except IndexError:
                    uliza = ''
               try:
                    dom = grab.doc.select(u'//td[contains(text(),"Состояние")]/following-sibling::td').text()
               except IndexError:
                    dom = ''
		    
               try:
                    tip = grab.doc.select(u'//p[contains(text(),"Средняя цена на похожие объекты -")]/b').text().split(', ')[0]
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//td[contains(text(),"Назначение")]/following-sibling::td').text()
               except IndexError:
                    naz =''
               try:
		    t1 =  grab.doc.select(u'//div[@id="price_table"]/following::p[1]').text()
		    if u'Добавлено:' in t1:
                         klass =  grab.doc.select(u'//div[@id="price_table"]/following::p[2]').text().split(', ')[2]
		    else:
		         klass = t1.split(', ')[2] 
                    
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//span[@itemprop="offers"]/b').text()
               except IndexError:
                    price =''
               try: 
                    plosh = grab.doc.select(u'//td[contains(text(),"Площадь")]/following-sibling::td').text().split('/')[0]+u' м2'
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//td[contains(text(),"Этаж")]/following-sibling::td').text().split('/')[0]
               except IndexError:
                    ohrana =''
               try:
                    gaz =  grab.doc.select(u'//td[contains(text(),"Материал")]/following-sibling::td').text()
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//td[contains(text(),"Этаж")]/following-sibling::td').text().split('/')[1]
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//td[contains(text(),"Высота потолков")]/following-sibling::td').text()
               except IndexError:
                    kanal =''
               try:
                    elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
               except DataNotFound:
                    elek =''
               try:
                    teplo =  grab.doc.select(u'//td[contains(text(),"За кв.м.")]/following-sibling::td[1]').text()+u' р.'
               except DataNotFound:
                    teplo =''
               #time.sleep(1)
	       try:
		    ln = []
		    for m in grab.doc.select(u'//h2[contains(text(),"Рекламная информация")]/following-sibling::p'):
		         urr = m.text()
		         ln.append(urr)
		    opis = "".join(ln)    
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//p[contains(text(),"Имя:")]/b').text()
	       except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//p[contains(text(),"Агентство")]/a').text()
               except IndexError:
                    comp = ''
               try:
                    data1 = grab.doc.select(u'//p[contains(text(),"Обновлено")]').text().split(u'Обновлено: ')[1] 
               except IndexError:   
                    data1 = ''
	       try: 
	            data = re.sub('[^\d\.]','',grab.doc.select(u'//p[contains(text(),"Добавлено")]').text().split(u' в ')[0])
	       except IndexError:
		    data= datetime.today().strftime('%d.%m.%Y')
	       
	       try:
                    phone = re.sub('[^\d\+\,]','',grab.doc.select(u'//p[contains(text(),"Тел.:")]').text())
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
	                   'klass': klass.replace(mesto,'').replace(u' район','').replace(punkt,''),
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
	       self.ws.write(self.result, 35, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 15, task.project['ulica'])
	       self.ws.write(self.result, 18, task.project['dom'])
	       self.ws.write(self.result, 8, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       self.ws.write(self.result, 3, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 13, task.project['ohrana'])
	       self.ws.write(self.result, 16, task.project['gaz'])
	       self.ws.write(self.result, 14, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 23, task.project['electr'])
	       self.ws.write(self.result, 36, task.project['teplo'])
	       self.ws.write(self.result, 25, task.project['opis'])
	       self.ws.write(self.result, 26, u'RealtyVision.ru')
	       self.ws.write_string(self.result, 27, task.project['url'])
	       self.ws.write(self.result, 28, task.project['phone'])
	       self.ws.write(self.result, 29, task.project['lico'])
	       self.ws.write(self.result, 30, task.project['company'])
	       self.ws.write(self.result, 31, task.project['data'])
	       self.ws.write(self.result, 32, task.project['data1'])
	       self.ws.write(self.result, 33, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 34, oper)
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print '***',i+1,'/',len(l),'***'
	      
	       print('*'*100)
	       self.result+= 1
	       
	      
	       
	       
	       
	       #if self.result > 20:
	            #self.stop()	       
	


     bot = Realtyvision_Com(thread_number=5, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=5, connect_timeout=5)
     bot.run()
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
          if oper == u'Продажа':
               i = 0
               l= open('links/com_arenda.txt').read().splitlines()
               dc = len(l)
               page = l[i]
               oper = u'Аренда'
          else:
               break
       
     
     
     
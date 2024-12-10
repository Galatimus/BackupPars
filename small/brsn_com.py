#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
import time
import os
import xlsxwriter
from datetime import datetime,timedelta

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'comm/0001-0016_00_C_001-0047_BRSNRU.xlsx')       

l= ['http://www.brsn.ru/comercheskay.html','http://www.brsn.ru/comercheskay/arenda-comercheskay.html']


class Brsn_Com(Spider):
    
    
    
    def prepare(self):
	 
	
	self.ws = workbook.add_worksheet()
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
	self.result= 1
	
            
            
            
              
    
    def task_generator(self):
        for line in l:#.read().splitlines():
            yield Task ('post',url=line.strip(),network_try_count=100)
	
	
	
    def task_page(self,grab,task):
        try:
            pg = grab.doc.select(u'//li[@class="pagination-next"]/a')
            u = grab.make_url_absolute(pg.attr('href'))
            yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
        except DataNotFound:
            print('*'*100)
            print '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!','NO PAGE NEXT','!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
            print('*'*100)
            logger.debug('%s taskq size' % self.task_queue.size())	
        
    def task_post(self,grab,task):
        #try:
            #num = re.sub("[^0-9]", "",grab.doc.select(u'//span[@class="objects_list_title_site_selected"]').text())
            ##print num
        #except DataNotFound:
            #num = ''
	for elem in grab.doc.select(u'//div[@class="photo-container"]/a'):
	    ur = grab.make_url_absolute(elem.attr('href'))  
	    print ur
	    yield Task('item', url=ur,refresh_cache=True)
	yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)
	            
  
	
        
        
        
    def task_item(self, grab, task):
	
        try:
            sub = u'Брянская область'#grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[0]
        except DataNotFound:
            sub = ''	
	try:
            ray = grab.doc.select(u'//div[@class="sttnmls_navigation text-center"]/a[contains(@href,"raion")]').text().replace(u'р-н ','')
        except DataNotFound:
            ray =''
        try:
            #if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text(),"район")]').exists()==True:
                #punkt= grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[2]
            #else:
            punkt= grab.doc.select(u'//div[@class="sttnmls_navigation text-center"]/a[contains(@href,"city")]').text()
        except IndexError:
            punkt = ''
        try:
            if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text(),"район")]').exists()==True:
                ter= grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[3].replace(u'улица','')
            else:
                ter= ''#grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[2]
        except IndexError:
            ter =''	    
	try:
	    #try:
                #uliza = grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text(),"улица")]').text()
	    #except DataNotFound:
	        #uliza = grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text(),"проспект")]').text()
            #except DataNotFound:
            uliza = grab.doc.select(u'//div[@class="sttnmls_navigation text-center"]/a[contains(@href,"street")]').text()
        except DataNotFound:
            uliza = '' 
	try:
	    dom = grab.doc.select(u'//h1[@class="offer-title"]').number()
	    #if re.sub(u'[^\d]','',d).isdigit()==True:
	        #dom = d.split(', ')[0]
	    #else:
	        #dom = ''
	except DataNotFound:
	    dom = ''
	    
	try:
            seg = grab.doc.select(u'//h1').text().split(' ')[1].replace(',','')
          #print oren
        except DataNotFound:
            seg = '' 
	    
	try:
            naz = grab.doc.select(u'//p[@class="text-justify"]/span').text().split(u'используется как ')[1].replace(',','')
          #print naz
        except IndexError:
	    naz = '' 
	    
        try:
            klass = grab.doc.rex_text(u'класса (.*?). ')
        except DataNotFound:
            klass = ''
	    
	try:
	    price = grab.doc.select(u'//div[@class="sttnmls_price"]/div[1]').text().replace(' ','').replace(u'a',u' р.')
	  #print price
	except DataNotFound:
	    price = ''
	    
	try:
            plosh = grab.doc.select(u'//p[@class="text-justify"]/span[contains(text(),"площадь:")]/b[1]').text()+u' м2'
          #print plosh
        except DataNotFound:
            plosh = '' 
	    
        try:
            et = grab.doc.rex_text(u', <b>(.*?)</b> этаж')
        except DataNotFound:
            et = ''
	    
        try:
            mat = grab.doc.select(u'//p[@class="text-justify"]/span').text().split(', ')[0]
        except DataNotFound:
            mat = ''
	    
	try:
            opis = grab.doc.select(u'//div[@id="dopinfo"]/p').text()
          #print opis
        except DataNotFound:
            opis = ''
	    
        try:
            phone = grab.doc.rex_text(u'href="tel:(.*?)">')
          #print phone
        except DataNotFound:
            phone = '' 
	    
        try:
            lico = grab.doc.select(u'//img[@class="thumbnail"][contains(@src,"components")]/following::b[1]').text()
        except IndexError:
            lico = ''
	    
	try:
            comp = grab.doc.select(u'//img[@class="thumbnail"][contains(@src,"agency")]/following::b[1]').text()
          #print comp
          
        except DataNotFound:
            comp = '' 
	try:
	    con = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
	        (u' мая ',u'.05.'),(u' июня ',u'.06.'),
	        (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
	        (u' января ',u'.01.'),(u' декабря ',u'.12.'),
	        (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
	        (u' февраля ',u'.02.'),(u' октября ',u'.10.')] 
	    dt1= grab.doc.select(u'//b[contains(text(),"Добавлено:")]/following-sibling::span[1]').text()
	    ohrana = reduce(lambda dt1, r: dt1.replace(r[0], r[1]), con, dt1).replace(' ','')
	except DataNotFound:
	    ohrana =''
	try:
	    gaz = re.sub(u'^.*(?=газ)','', grab.doc.select(u'//*[contains(text(), "газ")]').text())[:3].replace(u'газ',u'есть')
	except DataNotFound:
	    gaz =''
	try:
	    voda = re.sub(u'^.*(?=вод)','', grab.doc.select(u'//*[contains(text(), "вод")]').text())[:3].replace(u'вод',u'есть')
	except DataNotFound:
	    voda =''
	try:
	    kanal = re.sub(u'^.*(?=санузел)','', grab.doc.select(u'//*[contains(text(), "санузел")]').text())[:7].replace(u'санузел',u'есть')
	except DataNotFound:
	    kanal =''
	try:
	    elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
	except DataNotFound:
	    elek =''
	try:
	    teplo = grab.doc.select(u'//div[@class="sttnmls_navigation text-center"]').text().replace(u' → ',', ')
	except DataNotFound:
	    teplo =''
	    
	try:
            conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
	            (u' мая ',u'.05.'),(u' июня ',u'.06.'),
	            (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
	            (u' января ',u'.01.'),(u' декабря ',u'.12.'),
	            (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
	            (u' февраля ',u'.02.'),(u' октября ',u'.10.')] 
            dt= grab.doc.select(u'//b[contains(text(),"Обновлено:")]/following-sibling::span').text()#.split(', ')[0]
            data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(' ','')
		   #print data
        except DataNotFound:
            data = ''
	    
        try:
            oper = grab.doc.select(u'//h1').text().split(' ')[0].replace(u'Продаются',u'Продажа').replace(u'Сдаются',u'Аренда')
        except DataNotFound:
            oper = ''
	
	projects = {'url': task.url,
	            'sub': sub,
	            'ray': ray,
	            'punkt': punkt,
	            'teritor': ter,
	            'uliza': uliza,
	            'dom': dom,
	            'seg': seg,
	            'naznachenie': naz,
	            'klass': klass,
	            'cena': price,
	            'ploshad': plosh,
	            'et': et,
	            'mat': mat,
	            'opisanie': opis,
	            'phone':phone,
	            'company':comp,
	            'lico':lico,
	            'ohrana':ohrana,
                    'gaz': gaz,
                    'voda': voda,
                    'kanaliz': kanal,
                    'electr': elek,
                    'teplo': teplo,
	            'data':data,
	            'oper':oper
	            
	            }
	yield Task('write',project=projects,grab=grab)
	
    def task_write(self,grab,task):
	
	print('*'*50)
	print  task.project['sub']
	print  task.project['ray']
	print  task.project['punkt']
	print  task.project['teritor']
	print  task.project['uliza']
	print  task.project['dom']
	print  task.project['seg']
	print  task.project['naznachenie']
	print  task.project['klass']
	print  task.project['cena']
	print  task.project['ploshad']
	print  task.project['et']
	print  task.project['mat']
	print  task.project['opisanie']
	print  task.project['url']
	print  task.project['phone']
	print  task.project['lico']
	print  task.project['company']
	print  task.project['ohrana']
	print  task.project['gaz']
	print  task.project['voda']
	print  task.project['kanaliz']
	print  task.project['electr']
	print  task.project['teplo']
	print  task.project['data']
	print  task.project['oper']
	
	
	
	self.ws.write(self.result, 0, task.project['sub'])
	self.ws.write(self.result, 1, task.project['ray'])
	self.ws.write(self.result, 2, task.project['punkt'])
	self.ws.write(self.result, 3, task.project['teritor'])
	self.ws.write(self.result, 4, task.project['uliza'])
	self.ws.write(self.result, 5, task.project['dom'])
	#self.ws.write(self.result, 6, task.project['orentir'])
	self.ws.write(self.result, 7, task.project['seg'])
	#self.ws.write(self.result, 8, task.project['tip'])
	self.ws.write(self.result, 9, task.project['naznachenie'])
	self.ws.write(self.result, 10, task.project['klass'])
	self.ws.write(self.result, 11, task.project['cena'])
	self.ws.write(self.result, 12, task.project['ploshad'])	
	self.ws.write(self.result, 13, task.project['et'])
	#self.ws.write(self.result, 14, task.project['ets'])
	#self.ws.write(self.result, 15, task.project['god'])
	self.ws.write(self.result, 16, task.project['mat'])
	#self.ws.write(self.result, 17, task.project['potolok'])
	#self.ws.write(self.result, 18, task.project['sost'])
	self.ws.write(self.result, 31, task.project['ohrana'])
	self.ws.write(self.result, 20, task.project['gaz'])
	self.ws.write(self.result, 21, task.project['voda'])
	self.ws.write(self.result, 22, task.project['kanaliz'])
	self.ws.write(self.result, 23, task.project['electr'])
	self.ws.write(self.result, 35, task.project['teplo'])
	self.ws.write_string(self.result, 27, task.project['url'])
	self.ws.write(self.result, 28, task.project['phone'])
	self.ws.write(self.result, 29, task.project['lico'])
	self.ws.write(self.result, 30, task.project['company'])
	self.ws.write(self.result, 32, task.project['data'])
	self.ws.write(self.result, 25, task.project['opisanie'])
	self.ws.write(self.result, 26, u'Брянский сервер недвижимости')
	self.ws.write(self.result, 33, datetime.today().strftime('%d.%m.%Y'))
	self.ws.write(self.result, 34, task.project['oper'])
	print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	logger.debug('Tasks - %s' % self.task_queue.size()) 
	print('*'*50)
	
	self.result+= 1
	
	

        
    
            
            
            


        
        
 
       
bot = Brsn_Com(thread_number=1,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5000, connect_timeout=5000)
bot.run()
time.sleep(2)
workbook.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/small/dom43_zem.py")



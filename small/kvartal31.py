#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,GrabConnectionError 
import logging
import re
import time
import os
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'comm/0001-0081_00_C_001-0014_ANKV31.xlsx')       




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
	self.ws.write(0, 31, u"ДАТА_ОБНОВЛЕНИЯ")
	self.ws.write(0, 32, u"ДАТА_ПАРСИНГА")
	self.ws.write(0, 33, u"ОПЕРАЦИЯ")
	self.ws.write(0, 34, u"ЦЕНА_ЗА_М2")
	self.result= 1
	
           
            
            
              
    
    def task_generator(self):
        yield Task ('post',url=u'http://ankvartal31.ru/%D0%BA%D1%83%D0%BF%D0%B8%D1%82%D1%8C-%D0%BA%D0%BE%D0%BC%D0%BC%D0%B5%D1%80%D1%87%D0%B5%D1%81%D0%BA%D1%83%D1%8E-%D0%BD%D0%B5%D0%B4%D0%B2%D0%B8%D0%B6%D0%B8%D0%BC%D0%BE%D1%81%D1%82%D1%8C-%D0%B2-%D0%B1%D0%B5%D0%BB%D0%B3%D0%BE%D1%80%D0%BE%D0%B4%D0%B5',network_try_count=100)
	
	
	
   
        
    def task_post(self,grab,task):
        for elem in grab.doc.select(u'//div[@class="views-field-title"]/a'):
	    ur = grab.make_url_absolute(elem.attr('href'))  
	    print ur
	    yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	
	            
  
	
        
        
        
    def task_item(self, grab, task):
	
        try:
            sub = u'Белгородская область'#grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[0]
        except IndexError:
            sub = ''	
	try:
            ray = grab.doc.select(u'//div[@class="sttnmls_navigation text-center"]/a[contains(@href,"raion")]').text().replace(u'р-н ','')
        except IndexError:
            ray =''
        try:
            #if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text(),"район")]').exists()==True:
                #punkt= grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[2]
            #else:
            punkt= 'Белгород'#grab.doc.select(u'//div[@class="sttnmls_navigation text-center"]/a[contains(@href,"city")]').text()
        except IndexError:
            punkt = ''
        try:
            ter= grab.doc.select(u'//th[contains(text(),"Район города")]/following::div[3]').text()#.split(', ')[3].replace(u'улица','')
            
        except IndexError:
            ter =''	    
	try:
	    #try:
                #uliza = grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text(),"улица")]').text()
	    #except IndexError:
	        #uliza = grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text(),"проспект")]').text()
            #except IndexError:
            uliza = grab.doc.select(u'//th[contains(text(),"Адрес")]/following::div[3]').text().split(', ')[0]
        except IndexError:
            uliza = '' 
	try:
	    dom = grab.doc.select(u'//th[contains(text(),"Адрес")]/following::div[3]').text().split(', ')[1]
	    #if re.sub(u'[^\d]','',d).isdigit()==True:
	        #dom = d.split(', ')[0]
	    #else:
	        #dom = ''
	except IndexError:
	    dom = ''
	    
	try:
            seg = grab.doc.select(u'//span[@class="breadcrumb-separator"]/following-sibling::text()').text()#.split(' ')[1].replace(',','')
          #print oren
        except IndexError:
            seg = '' 
	    
	try:
            naz = grab.doc.select(u'//th[contains(text(),"Категория недвижимости")]/following::div[3]').text()#.split(u'используется как ')[1].replace(',','')
          #print naz
        except IndexError:
	    naz = '' 
	    
        try:
            klass = grab.doc.select(u'//th[contains(text(),"Микрорайон")]/following::div[3]').text()
        except IndexError:
            klass = ''
	    
	try:
	    price = grab.doc.select(u'//div[@class="field field-name-field-mt-price field-type-number-integer field-label-hidden ffstoim"]/div/div').text()#.replace(' ','').replace(u'a',u' р.')
	  #print price
	except IndexError:
	    price = ''
	    
	try:
            plosh = grab.doc.select(u'//th[contains(text(),"Общая площадь")]/following::div[3]').text()#+u' м2'
          #print plosh
        except IndexError:
            plosh = '' 
	    
        try:
            et = grab.doc.select(u'//th[contains(text(),"Этаж")]/following-sibling::td').text()
        except IndexError:
            et = ''
	    
        try:
            mat = grab.doc.select(u'//th[contains(text(),"Тип дома")]/following::div[3]').text()#.split(', ')[0]
        except IndexError:
            mat = ''
	    
	try:
	        try:
                    opis = grab.doc.select(u'//div[@class="field-item even"]/h2/following-sibling::p').text()
                except IndexError:
		    opis = grab.doc.select(u'//div[@class="field-item even"]/p').text()
        except IndexError:
            opis = ''
	    
        try:
            phone = grab.doc.select(u'//i[@class="fa fa-phone"]/following-sibling::text()').text()
          #print phone
        except IndexError:
            phone = '' 
	    
        try:
            lico = grab.doc.select(u'//i[@class="fa fa-user"]/following-sibling::a').text()
        except IndexError:
            lico = ''
	    
	try:
            comp = u'Квартал'#grab.doc.select(u'//img[@class="thumbnail"][contains(@src,"agency")]/following::b[1]').text()
          #print comp
          
        except IndexError:
            comp = '' 
	try:
	    ohrana = grab.doc.select(u'//th[contains(text(),"Этажность")]/following-sibling::td').text()
	except IndexError:
	    ohrana =''
	try:
	    gaz = grab.doc.select(u'//th[contains(text(),"Год постройки")]/following-sibling::td').text()
	except IndexError:
	    gaz =''
	try:
	    voda = grab.doc.select(u'//th[contains(text(),"Высота потолков")]/following::div[3]').text()
	except IndexError:
	    voda =''
	try:
	    kanal = grab.doc.select(u'//th[contains(text(),"Ремонт")]/following::div[3]').text()
	except IndexError:
	    kanal =''
	try:
	    elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
	except IndexError:
	    elek =''
	try:
	    teplo = grab.doc.select(u'//th[contains(text(),"Стоимость кв.м")]/following::div[3]').text()+' р.'
	except IndexError:
	    teplo =''
	    
	try:
            data= grab.doc.select(u'//span[@class="submitted-by"]').text().split(': ')[1][:11]
        except IndexError:
            data = ''
	    
        try:
            oper = u'Продажа'
        except IndexError:
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
	print  task.project['ohrana']
	print  task.project['mat']
	print  task.project['opisanie']
	print  task.project['url']
	print  task.project['phone']
	print  task.project['lico']
	print  task.project['company']
	
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
	self.ws.write(self.result, 8, task.project['seg'])
	#self.ws.write(self.result, 8, task.project['tip'])
	self.ws.write(self.result, 9, task.project['naznachenie'])
	self.ws.write(self.result, 6, task.project['klass'])
	self.ws.write(self.result, 11, task.project['cena'])
	self.ws.write(self.result, 12, task.project['ploshad'])	
	self.ws.write(self.result, 13, task.project['et'])
	#self.ws.write(self.result, 14, task.project['ets'])
	#self.ws.write(self.result, 15, task.project['god'])
	self.ws.write(self.result, 16, task.project['mat'])
	#self.ws.write(self.result, 17, task.project['potolok'])
	#self.ws.write(self.result, 18, task.project['sost'])
	self.ws.write(self.result, 14, task.project['ohrana'])
	self.ws.write(self.result, 15, task.project['gaz'])
	self.ws.write(self.result, 17, task.project['voda'])
	self.ws.write(self.result, 18, task.project['kanaliz'])
	self.ws.write(self.result, 23, task.project['electr'])
	self.ws.write(self.result, 34, task.project['teplo'])
	self.ws.write_string(self.result, 27, task.project['url'])
	self.ws.write(self.result, 28, task.project['phone'])
	self.ws.write(self.result, 29, task.project['lico'])
	self.ws.write(self.result, 30, task.project['company'])
	self.ws.write(self.result, 31, task.project['data'])
	self.ws.write(self.result, 25, task.project['opisanie'])
	self.ws.write(self.result, 26, u'АН "Квартал" г. Белгород')
	self.ws.write(self.result, 32, datetime.today().strftime('%d.%m.%Y'))
	self.ws.write(self.result, 33, task.project['oper'])
	print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	logger.debug('Tasks - %s' % self.task_queue.size()) 
	print('*'*50)
	
	self.result+= 1
	
	

        
    
            
            
            


        
        
 
       
bot = Brsn_Com(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')

bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
try:
    command = 'mount -a'
    os.system('echo %s|sudo -S %s' % ('1122', command))
    time.sleep(2)
    workbook.close()
    print('Done')
except IOError:
    time.sleep(30)
    os.system('echo %s|sudo -S %s' % ('1122', 'mount -a'))
    time.sleep(10)
    workbook.close()
    print('Done!')
time.sleep(5)
os.system("/home/oleg/pars/small/zemru.py")
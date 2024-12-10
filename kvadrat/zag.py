#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import re
import time
import xlsxwriter
from datetime import datetime

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

g = Grab(timeout=2000, connect_timeout=2000)

i = 0


l= ['http://kvadrat22.ru/mob_sellzagbank-1000-1.html',
    'http://kvadrat24.ru/mob_sellzagbank-1000-1.html',
    'http://kvadrat54.ru/mob_sellzagbank-1000-1.html',
    'http://kvadrat64.ru/mob_sellzagbank-1000-1.html',
    'http://kvadrat66.ru/mob_sellzagbank-1000-1.html',
    'http://kvadrat72.ru/mob_sellzagbank-1000-1.html',
    'http://kvadrat74.ru/mob_sellzagbank-1000-1.html',
    'http://n30.ru/mob_sellzagbank-1000-1.html',
    'http://kemdom.ru/mob_sellzagbank-1000-1.html',
    'http://n002.ru/mob_sellzagbank-1000-1.html',
    'http://kazan-n.ru/mob_sellzagbank-1000-1.html',
    'http://nd27.ru/mob_sellzagbank-1000-1.html',
    'http://kvadrat22.ru/mob_givezagbank-1000-1.html',
    'http://kvadrat24.ru/mob_givezagbank-1000-1.html',
    'http://kvadrat54.ru/mob_givezagbank-1000-1.html',
    'http://kvadrat64.ru/mob_givezagbank-1000-1.html',
    'http://kvadrat66.ru/mob_givezagbank-1000-1.html',
    'http://kvadrat72.ru/mob_givezagbank-1000-1.html',
    'http://kvadrat74.ru/mob_givezagbank-1000-1.html',
    'http://n30.ru/mob_givezagbank-1000-1.html',
    'http://kemdom.ru/mob_givezagbank-1000-1.html',
    'http://n002.ru/mob_givezagbank-1000-1.html',
    'http://kazan-n.ru/mob_givezagbank-1000-1.html',
    'http://nd27.ru/mob_givezagbank-1000-1.html']

page = l[i]
while True:
     
     class Kvadrat_Zag(Spider):
	  def prepare(self):
	       self.f = page
	       while True:
		   try:
		       time.sleep(1)
		       g.go(self.f)
		       conv = [(u'Хабаровска',u'Хабаровский край'),(u'Барнаула',u'Алтайский край'),
			       (u'Красноярска',u'Красноярский край'),(u'Саратова',u'Саратовская область'),
			       (u'Новосибирска',u'Новосибирская область'),(u'Екатеринбурга',u'Свердловская область'),
			       (u'Тюмени',u'Тюменская область'),(u'Челябинска',u'Челябинская область'),
			       (u'Астрахани',u'Астраханская область'),(u'Кемерово',u'Кемеровская область'),
			       (u'Уфы',u'Башкортостан'),(u'Казани',u'Татарстан')]        
		       dt = g.doc.select(u'//a[@class="amenu"][contains(text(),"Недвижимость")]').text().replace(u'Недвижимость','') 
		       self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
		       print self.sub
		       break
		   except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
		       print g.config['proxy'],'Change proxy'
		       g.change_proxy()
		       continue
	       self.workbook = xlsxwriter.Workbook(u'zag/Kvadrat_%s' % bot.sub +str(i+1)+ u'_Загород.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Kvadrat_Загород')
	       self.ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
	       self.ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
	       self.ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
	       self.ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
	       self.ws.write(0, 4, u"УЛИЦА")
	       self.ws.write(0, 5, u"ДОМ")
	       self.ws.write(0, 6, u"ОРИЕНТИР")
	       self.ws.write(0, 7, u"ТРАССА")
	       self.ws.write(0, 8, u"УДАЛЕННОСТЬ")
	       self.ws.write(0, 9, u"КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
	       self.ws.write(0, 10, u"ТИП_ОБЪЕКТА")
	       self.ws.write(0, 11, u"ОПЕРАЦИЯ")
	       self.ws.write(0, 12, u"СТОИМОСТЬ")
	       self.ws.write(0, 13, u"ЦЕНА_М2")
	       self.ws.write(0, 14, u"ПЛОЩАДЬ_ОБЩАЯ")
	       self.ws.write(0, 15, u"КОЛИЧЕСТВО_КОМНАТ")
	       self.ws.write(0, 16, u"ЭТАЖНОСТЬ")
	       self.ws.write(0, 17, u"МАТЕРИАЛ_СТЕН")
	       self.ws.write(0, 18, u"ГОД_ПОСТРОЙКИ")
	       self.ws.write(0, 19, u"ПЛОЩАДЬ_УЧАСТКА")
	       self.ws.write(0, 20, u"ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
	       self.ws.write(0, 21, u"ГАЗОСНАБЖЕНИЕ")
	       self.ws.write(0, 22, u"ВОДОСНАБЖЕНИЕ")
	       self.ws.write(0, 23, u"КАНАЛИЗАЦИЯ")
	       self.ws.write(0, 24, u"ЭЛЕКТРОСНАБЖЕНИЕ")
	       self.ws.write(0, 25, u"ТЕПЛОСНАБЖЕНИЕ")
	       self.ws.write(0, 26, u"ЛЕС")
	       self.ws.write(0, 27, u"ВОДОЕМ")
	       self.ws.write(0, 28, u"БЕЗОПАСНОСТЬ")
	       self.ws.write(0, 29, u"ОПИСАНИЕ")
	       self.ws.write(0, 30, u"ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 31, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
	       self.ws.write(0, 32, u"ТЕЛЕФОН")
	       self.ws.write(0, 33, u"КОНТАКТНОЕ_ЛИЦО")
	       self.ws.write(0, 34, u"КОМПАНИЯ")
	       self.ws.write(0, 35, u"ДАТА_РАЗМЕЩЕНИЯ")
	       self.ws.write(0, 36, u"ДАТА_ПАРСИНГА")
	       self.ws.write(0, 37, u"ВИД_РАЗРЕШЕННОГО_ИСПОЛЬЗОВАНИЯ")
		    
		    
	       self.result= 1
	     
		    
	 
	  def task_generator(self):
	       yield Task ('post',url = page,network_try_count=1000)
		 
		 
	  def task_page(self,grab,task):
	       try:         
		    pg = grab.doc.select(u'//div[@class="dphase"]/following-sibling::a[1]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,network_try_count=1000)
	       except DataNotFound:
		    print('*'*50)
		    print '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!','NO PAGE NEXT','!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
		    print('*'*50)
		    logger.debug('%s taskq size' % self.task_queue.size())             
	     
	     
		 
		 
	  def task_post(self,grab,task):
		       
	       for elem in grab.doc.select(u'//a[@class="site3"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))  
		    #print ur
		    yield Task('item', url=ur,network_try_count=1000)
	       yield Task("page", grab=grab,network_try_count=1000)
	     
	     
	  def task_item(self, grab, task):
	       try:   
	            punkt= re.sub(r'\(.*?$', '',grab.doc.select(u'//td[@class="hh"]').text().split(', ')[1])
	       except IndexError:
		    punkt = ''
	       try:
		    ter= re.sub(r'\(.*?$', '',grab.doc.select(u'//td[@class="hh"]').text().split(', ')[2])
	       except IndexError:
		    ter =''
	       try:
		    uliza=grab.doc.rex_text(u'адрес:(.*?)</span>')
	       except IndexError:
		    uliza = ''
	       try:
		    dom = re.sub('[^0-9a-f]', '',re.sub(r'\(.*?$', '',grab.doc.select(u'//td[@class="hh"]').text().split(', ')[3]).replace(u' на карте',''))[:2]
	       except IndexError:
		    dom = ''
		    
	       try:
		    udal = grab.doc.select(u'//td[@class="hh"]/span[1]').text().split(')')[0].replace('(','')
	       except IndexError:
		    udal = ''
	       try:
		    tip_ob = grab.doc.select(u'//td[@class="hh"]').text().split(' ')[0]
	       except IndexError:
		    tip_ob = ''
               try:
                    oper = grab.doc.select(u'//div[@class="a"]').text().split(' ')[0]
               except IndexError:
                    oper = ''		    
		    
	       try:
		    price = grab.doc.select(u'//td[@class="thprice"]').text()
	       except IndexError:
		    price = ''
		    
	       
		    
	       try:
		    plosh = grab.doc.rex_text(u'Площадь дома:<br><span class=d>(.*?)</span>').replace('&sup','').replace(';','')
	       except IndexError:
		    plosh = ''
		    
	       try:
		    etash = re.sub('[^\d]', u'',grab.doc.rex_text(u'Этажность дома:<br><span class=d>(.*?)</span>'))
	       except IndexError:
		    etash = ''
		    
	       try:
		    plosh_uch = grab.doc.rex_text(u'Площадь участка:<br><span class=d>(.*?)</span>')
	       except IndexError:
		    plosh_uch = ''
	       
	       try:
		    mat = grab.doc.rex_text(u'Материал дома:<br><span class=d>(.*?)</span>')
	       except IndexError:
		    mat = ''	  
		    
	       try:
		    vid = grab.doc.rex_text(u'Год постройки:<br><span class=d>(.*?)</span>')
	       except IndexError:
		    vid = ''
               try:
	            postr = grab.doc.rex_text(u'Постройки, посадки:<br><span class=d>(.*?)</span>')
	       except IndexError:
                    postr = ''
		    
		    
	       try:
		    ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
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
		    teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
	       except DataNotFound:
		    teplo =''
	       try:
		    les = re.sub(u'^.*(?=лес)','', grab.doc.select(u'//*[contains(text(), "лес")]').text())[:3].replace(u'лес',u'есть')
	       #gazz = gaz.replace('True',u'есть')
	       except DataNotFound:
		    les =''
		 
	       try:
		    vodoem = re.sub(u'^.*(?=озер)','', grab.doc.select(u'//*[contains(text(), "озер")]').text())[:4].replace(u'озер',u'есть')
	       #gazz = gaz.replace('True',u'есть')
	       except DataNotFound:
		    vodoem =''	  
		    
	                    
		   
			 
	       try:
		    opis = grab.doc.select(u'//div[contains(text(), "Дополнительная информация:")]/span').text() 
	       except IndexError:
		    opis = ''
		    
	       try:
		    phone = re.sub('[^\d]','',grab.doc.select(u'//div[@class="divdec"]').text().split(u'телефон:')[1])[:11] 
	       except IndexError:
		    phone = ''
		    
	       try:
		    lico = grab.doc.rex_text(u'Персона для контактов:<br><span class=d>(.*?)</span>')
	       except IndexError:
		    lico = ''
		    
	       
		    
	       try:
		    data = grab.doc.select(u'//td[@class="tdate"]').text().split(u'создано ')[0].replace(u'обновлено ','').replace('-','.')
	       except IndexError:
		    data = ''
		    
               try:
	            istoch = grab.doc.select(u'//a[@class="amenu"][contains(text(),"Недвижимость")]').text().replace(u'Недвижимость',u'Недвижимость ')
               except IndexError:
                    istoch = ''		    
			 
	       
							
		    
	       projects = {'url': task.url,
		           'sub': self.sub,
		           'punkt': punkt,
		           'teritor': ter,
		           'ulica': uliza,
		            'dom': dom,
		           'udal': udal,
		           'object': tip_ob,
		           'cena': price,
		           'plosh':plosh,
		           'etach': etash,
		           'plouh': plosh_uch,
		           'mat': mat,
		           'vid': vid,
	                   'stroenia': postr,
		           'ohrana':ohrana,
		           'gaz': gaz,
		           'voda': voda,
		           'kanaliz': kanal,
		           'electr': elek,
		           'teplo': teplo,
		           'les': les,
		           'vodoem':vodoem,	              
		           'opis':opis,
		           'phone':phone,
		           'lico':lico,
		           'data':data,
	                   'istochnik': istoch,
		           'oper':oper
		           }
	       
	       yield Task('write',project=projects,grab=grab,refresh_cache=True)
		 
	  def task_write(self,grab,task):
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['udal']
	       print  task.project['object']
	       print  task.project['cena']
	       print  task.project['plosh']
	       print  task.project['etach']
	       print  task.project['plouh']
	       print  task.project['mat']
	       print  task.project['vid']
	       print  task.project['stroenia']
	       print  task.project['ohrana']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['teplo']
	       print  task.project['les']
	       print  task.project['vodoem']	  
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['data']
	       print  task.project['oper']
	       print  task.project['istochnik']
	       
	       #global result
	       self.ws.write(self.result, 0, task.project['sub'])
	       #self.ws.write(self.result, 1, task.project['rayon'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 6, task.project['ulica'])
	       #self.ws.write(self.result, 7, task.project['trassa'])
	       self.ws.write(self.result, 8, task.project['udal'])
	       self.ws.write(self.result, 11, task.project['oper'])
	       self.ws.write(self.result, 10, task.project['object'])
	       self.ws.write(self.result, 12, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 21, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['etach'])
	       self.ws.write(self.result, 17, task.project['mat'])
	       self.ws.write(self.result, 23, task.project['kanaliz'])
	       self.ws.write(self.result, 24, task.project['electr'])
	       self.ws.write(self.result, 19, task.project['plouh'])
	       self.ws.write(self.result, 28, task.project['ohrana'])
	       self.ws.write(self.result, 22, task.project['voda'])	  
	       self.ws.write(self.result, 25, task.project['teplo'])
	       self.ws.write(self.result, 26, task.project['les'])
	       self.ws.write(self.result, 27, task.project['vodoem'])
	       self.ws.write(self.result, 29, task.project['opis'])
	       self.ws.write(self.result, 18, task.project['vid'])
	       self.ws.write(self.result, 30, task.project['istochnik'])
	       self.ws.write_string(self.result, 31, task.project['url'])
	       self.ws.write(self.result, 32, task.project['phone'])
	       self.ws.write(self.result, 33, task.project['lico'])
	       self.ws.write(self.result, 20, task.project['stroenia'])
	       self.ws.write(self.result, 35, task.project['data'])
	       self.ws.write(self.result, 36, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       #print task.sub
	       
	       print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       #print oper
	       print('*'*50)	       
	       self.result+= 1
		    
		    
		    
	       #if self.result > 10:
		    #self.stop()
     
	  
     bot = Kvadrat_Zag(thread_number=2,network_try_limit=100000)
     bot.load_proxylist('/home/oleg/Proxy/tipa.txt','text_file')
     bot.create_grab_instance(timeout=5000, connect_timeout=5000)
     bot.run()
     print bot.sub
     print(u'Сохранение...')
     print(u'Спим 2 сек...')
     time.sleep(2) 
     bot.workbook.close()
     print('Done!')
     i=i+1
     try:
	  page = l[i]
     except IndexError:
	  break 
     
     
     
     
     
     
     

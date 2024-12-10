#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import os
from grab import Grab
import logging
import time
import xlsxwriter
from datetime import datetime
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



i = 0
l= ['http://kvadrat22.ru/mob_selllandbank-1000-1.html',
    'http://kvadrat24.ru/mob_selllandbank-1000-1.html',
    'http://kvadrat54.ru/mob_selllandbank-1000-1.html',
    'http://kvadrat64.ru/mob_selllandbank-1000-1.html',
    'http://kvadrat66.ru/mob_selllandbank-1000-1.html',
    'http://kvadrat72.ru/mob_selllandbank-1000-1.html',
    'http://kvadrat74.ru/mob_selllandbank-1000-1.html',
    'http://n30.ru/mob_selllandbank-1000-1.html',
    'http://kemdom.ru/mob_selllandbank-1000-1.html',
    'http://n002.ru/mob_selllandbank-1000-1.html',
    'http://kazan-n.ru/mob_selllandbank-1000-1.html',
    'http://nd27.ru/mob_selllandbank-1000-1.html',
    #'http://nd23.ru/mob_selllandbank-1000-1.html',
    'http://kvadrat22.ru/mob_givelandbank-1000-1.html',
    'http://kvadrat24.ru/mob_givelandbank-1000-1.html',
    'http://kvadrat54.ru/mob_givelandbank-1000-1.html',
    'http://kvadrat64.ru/mob_givelandbank-1000-1.html',
    'http://kvadrat66.ru/mob_givelandbank-1000-1.html',
    'http://kvadrat72.ru/mob_givelandbank-1000-1.html',
    'http://kvadrat74.ru/mob_givelandbank-1000-1.html',
    'http://n30.ru/mob_givelandbank-1000-1.html',
    'http://kemdom.ru/mob_givelandbank-1000-1.html',
    'http://n002.ru/mob_givelandbank-1000-1.html',
    #'http://nd23.ru/mob_givelandbank-1000-1.html',
    'http://kazan-n.ru/mob_givelandbank-1000-1.html',
    'http://nd27.ru/mob_givelandbank-1000-1.html']

page = l[i]
while True:
     class Kvadrat_Zem(Spider): 
	  def prepare(self):
	       self.f = page
	       while True:
		    try:
			 time.sleep(1)
			 g = Grab(timeout=10, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
			 g.go(self.f)
			 conv = [(u'Хабаровска',u'Хабаровский край'),(u'Барнаула',u'Алтайский край'),
			           (u'Красноярска',u'Красноярский край'),(u'Саратова',u'Саратовская область'),
			           (u'Новосибирска',u'Новосибирская область'),(u'Екатеринбурга',u'Свердловская область'),
			           (u'Тюмени',u'Тюменская область'),(u'Челябинска',u'Челябинская область'),
			           (u'Астрахани',u'Астраханская область'),(u'Кемерово',u'Кемеровская область'),
			           (u'Уфы',u'Башкортостан'),(u'Казани',u'Татарстан'),(u'Краснодара',u'Краснодарский край')]        
			 dt = g.doc.select(u'//a[@class="amenu"][contains(text(),"Недвижимость")]').text().replace(u'Недвижимость','') 
			 self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
			 print self.sub
			 del g
			 break
		    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       self.workbook = xlsxwriter.Workbook(u'zem/Kvadrat_%s' % bot.sub +str(i+1)+ u'_Земля.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Kvadrat_ЗЕМЛЯ')
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
	       
	       self.result= 1
	       
          def task_generator(self):
	       yield Task ('post',url = page,network_try_count=100)
		    
	  def task_post(self,grab,task):
	       for elem in grab.doc.select(u'//td[@class="tdecprm"]/a[contains(@href,"mob")]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item',url=ur,network_try_count=100)
	       yield Task("page", grab=grab,network_try_count=100,use_proxylist=False)
	       
	  def task_page(self,grab,task): 
	       try:
	            pg = grab.doc.select(u'//div[@class="dphase"]/following-sibling::a[1]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u)
	       except DataNotFound:
	            print('*'*100)
	            print '!!!','NO PAGE NEXT','!!!'
	            print('*'*100)
	            logger.debug('%s taskq size' % self.task_queue.size())
	       
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
                    uli = re.sub(r'\(.*?$', '',grab.doc.select(u'//td[@class="hh"]').text().split(', ')[3]).replace(u' на карте','')
                    uliza=re.sub(r'[0-9a-f]', '',uli)
               except IndexError:
                    uliza = ''
               try:
                    dom = re.sub('[^0-9a-f]', '',re.sub(r'\(.*?$', '',grab.doc.select(u'//td[@class="hh"]').text().split(', ')[3]).replace(u' на карте',''))[:2]
               except IndexError:
                    dom = ''
               try:
                    orent = grab.doc.select(u'//td[@class="hh"]/span[2]').text().replace(u'(','').replace(u')','')
               except IndexError:
                    orent = ''
               try:
                    udal = grab.doc.select(u'//td[@class="hh"]/span[1]').text().replace(u'(','').replace(u')','')
               except IndexError:
                    udal = ''
               try:
                    oper = u'Продажа'
               except IndexError:
                    oper = ''
               try:
                    price = grab.doc.select(u'//td[@class="thprice"]').text()
               except IndexError:
                    price = ''
               try:
                    price_m = grab.doc.rex_text(u'Цена за м&sup2;:<br>(.*?)за')
               except IndexError:
		    price_m = ''
               try:
                    plosh = grab.doc.select(u'//div[@class="tddec"]').text().split(':')[1].replace(u'Тип земли','').replace(u'Коммуникации','')
               except IndexError:
                    plosh = ''
               try:
                    vid =grab.doc.rex_text(u'Тип земли:<br><span class=d>(.*?)</span>')
               except IndexError:
                    vid = ''
               try:
                    gaz = grab.doc.select(u'//td[@class="hh"]').text().split(u'соток, ')[1].replace(u' на карте','')
               except IndexError:
                    gaz =''
               try:
                    voda = grab.doc.rex_text(u'создано (.*?)</td>').replace('-','.')
               except IndexError:
                    voda =''
               try:
                    kanal = re.sub(u'^.*(?=канализация)','',grab.doc.select(u'//*[contains(text(), "канализация")]').text())[:11].replace(u'канализация',u'есть')
               except IndexError:
                    kanal =''
               try:
                    elek = re.sub(u'^.*(?=электричество)','',grab.doc.select(u'//*[contains(text(), "электричество")]').text())[:13].replace(u'электричество',u'есть')
               except IndexError:
                    elek =''
               try:
                    teplo = re.sub(u'^.*(?=отопление)','',grab.doc.select(u'//*[contains(text(), "отопление")]').text())[:9].replace(u'отопление',u'есть')
               except IndexError:
                    teplo =''
               try:
                    opis = grab.doc.select(u'//div[contains(text(), "Дополнительная информация:")]/span').text()
               except IndexError:
                    opis = ''
               try:
                    istoch = grab.doc.select(u'//a[@class="amenu"][contains(text(),"Недвижимость")]').text().replace(u'Недвижимость',u'Недвижимость ')
               except IndexError:
                    istoch = ''
               try:
		    tip = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[0])
		    user = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[1])
		    pkey = re.sub('[^\d]','',grab.doc.select(u'//span[@class="showphone"]').attr('onclick').split(',')[2])
		    link = task.url.split(u'mob')[0]#+u'ru'
		    url_ph = link+'showphone.php?tip='+tip+'&id='+user+'&from='+pkey
		    g2 = grab.clone()
		    g2.go(url_ph)
		    phone = re.sub('[^\d\,]','',re.findall('innerHTML="(.*?)"',g2.doc.body)[0])
		    del g2
		    ##print phone
	       except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
	            phone = ''
               try:
                    lico = re.sub(u'[^а-я^А-Я\ ]','',grab.doc.rex_text(u'посредник(.+?)</span>')) 
               except IndexError:
                    lico = ''
               try:
                    data = grab.doc.rex_text(u'обновлено (.*?)<br>').replace('-','.')
               except IndexError:
                    data = ''
   
              
              
              
              
              
              
               projects = {'sub': self.sub,
	                   'url': task.url,
	                      'punkt': punkt,
	                      'teritor': ter,
	                      'ulica': uliza,
	                      'dom': dom,
	                      'istochnik': istoch,
	                      'udal': udal,
	                      'cena': price,
	                      'plosh':plosh,
	                      'cena1':price_m,
	                      'vid': vid,
	                      'gaz': gaz,
	                      'voda': voda,
	                      'kanaliz': kanal,
	                      'electr': elek,
	                      'teplo': teplo,
	                      'opis':opis,
	                      'phone':phone,
	                      'lico':lico,
	                      'orentir':orent,
	                      'data':data,
	                      'oper':oper}
	       
	       
	       
	       
	       yield Task('write',project=projects,grab=grab)
	       
	  def task_write(self,grab,task):
	       print('*'*50)
	       print  task.project['sub']
	       print  task.project['punkt']
	       print  task.project['teritor']
	       print  task.project['ulica']
	       print  task.project['dom']
	       print  task.project['istochnik']
	       print  task.project['udal']
	       print  task.project['cena']
	       print  task.project['cena1']
	       print  task.project['plosh']
	       print  task.project['vid']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['teplo']
	       print  task.project['opis']
	       print task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['orentir']
	       print  task.project['data']
	       print  task.project['voda']
	       print  task.project['gaz']
	      

	       
	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 11, task.project['cena1'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 3, task.project['teritor'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 23, task.project['istochnik'])
	       self.ws.write(self.result, 8, task.project['udal'])
	       self.ws.write(self.result, 9, task.project['oper'])
	       self.ws.write_string(self.result, 10, task.project['cena'])
	       self.ws.write(self.result, 12, task.project['plosh'])
	       self.ws.write(self.result, 14, task.project['vid'])
	       self.ws.write(self.result, 31, task.project['gaz'])
	       self.ws.write(self.result, 28, task.project['voda'])
	       self.ws.write(self.result, 17, task.project['kanaliz'])
	       self.ws.write(self.result, 18, task.project['electr'])
	       self.ws.write(self.result, 19, task.project['teplo'])
	       self.ws.write(self.result, 6, task.project['orentir'])
	       self.ws.write(self.result, 22, task.project['opis'])
	       #self.ws.write(self.result, 23, u'MIRKVARTIR.RU')
	       self.ws.write_string(self.result, 24, task.project['url'])
	       self.ws.write(self.result, 25, task.project['phone'])
	       self.ws.write(self.result, 26, task.project['lico'])
	       #self.ws.write(self.result, 27, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))
	       print('*'*50)
	       print 'Ready - '+str(self.result)
	       logger.debug('Tasks - %s' % self.task_queue.size()) 
	       print '***',i+1,'/',len(l),'***'
	       print  task.project['oper']
               print('*'*100)
               self.result+= 1
	       
	       #if self.result > 20:
		    #self.stop()
	       
	

     bot = Kvadrat_Zem(thread_number=5,network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
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
	  break
     
time.sleep(5)
os.system("/home/oleg/pars/kvadrat/com.py")
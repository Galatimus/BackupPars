#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError
from grab import Grab
import logging
import random
import os
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







l= open('comp.txt').read().splitlines()
oper = u'Продажа'



for i in range(40):
     print '********************************************',i+1,'/',40,'*******************************************'

     class move_Com(Spider):
	  def prepare(self):
	       self.f = l[i]
               for p in range(1,31):
                    try:
                         time.sleep(1)
			 g = Grab(timeout=20, connect_timeout=50)
			 g.proxylist.load_file(path='../ivan.txt',proxy_type='http')
                         g.go(self.f)
			 self.sub  = g.doc.select(u'//div[@class="place"]/a').text()
			 self.pag = int(re.sub('[^\d]','',g.doc.select(u'//li[@class="text last-none"]/a').attr('href')))
			 self.num = self.pag*20
			 print self.sub
                         print self.pag
			 print self.num
			 del g
                         break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy',str(p),'/',str(20)
			 g.change_proxy()
			 del g
			 continue
	       else:
	            self.sub = ''
	            self.pag = 1
	            self.stop()	  

	       self.workbook = xlsxwriter.Workbook(u'com/Gde_'+oper+'_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet()
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
                    yield Task ('post',url=self.f+'?page=%d'%x,refresh_cache=True,network_try_count=50)


	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//div[@class="title"]/a'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=50)
	       #yield Task("page", grab=grab,refresh_cache=True,network_try_count=50)

	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//li[@class="text control forward last-none"]/a[contains(text(),"вперёд")]')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=50)
	       except IndexError:
		    print('*'*100)
		    print '!!!','NO PAGE NEXT','!!!'
		    print('*'*100)

	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//span[@class="address"]').text()
	       except IndexError:
	            mesto =''

	       try:
	            punkt= grab.doc.select(u'//a[contains(@href,"city")]').text()
	       except IndexError:
	            punkt = ''

               try:
                    tip = grab.doc.select(u'//td[contains(text(),"Вид недвижимости")]/following-sibling::td').text()
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//h1').text().split(',')[0].replace(u'Продам','').replace(u'Продажа','').replace(u'Продается','').replace(u'Продаю','').replace(u'Продаются','')
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//span[contains(text(),"Этаж")]/following::div[2]').text()
               except IndexError:
                    klass = ''
               try:

                    price = grab.doc.select(u'//span[@class="price"]').text()
               except IndexError:
                    price =''
               try:
                    plosh = grab.doc.select(u'//dt[contains(text(),"Площадь")]/following-sibling::dd').text()
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
                    voda =  grab.doc.select(u'//title').text()
               except IndexError:
                    voda =''
               try:
                    d1 = grab.doc.select(u'//div[@class="date"][2]').text().replace(u'Дата подачи: ','').split(u'г.')[0]
	            kanal = reduce(lambda d1, r: d1.replace(r[0], r[1]), self.conv, d1)
               except IndexError:
                    kanal =''
               try:
                    elek = re.sub(u'^.*(?=лектричество)','', grab.doc.select(u'//*[contains(text(), "лектричество")]').text())[:12].replace(u'лектричество',u'есть')
               except DataNotFound:
                    elek =''
               try:
                    teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
               except DataNotFound:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//p[@itemprop="description"]').text()
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//a[@class="user"]').text()
	       except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//a[@class="user"]').text().split(', ')[0]
               except IndexError:
                    comp = ''

	       try:
		    data = grab.doc.select(u'//div[@class="date"]').text().replace(u'с ','')
	       except IndexError:
		    data=''


	       #try:
                    #ad_phone = grab.doc.select(u'//div[@itemprop="telephone"]').attr('data-url')
		    #link = grab.make_url_absolute(ad_phone)
		    #headers ={'Accept': '*/*',
			      #'Accept-Encoding': 'gzip, deflate, br',
			      #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
			      ##'Cookie': 'sessid='+ ad_phone,
			      #'Host': 'gde.ru',
			      #'Referer': task.url,
			      #'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0',
			      #'X-Requested-With' : 'XMLHttpRequest'}
		    #g2 = grab.clone(headers=headers,proxy_auto_change=True)
		    #g2.request(headers=headers,url=link)
		    #print  g2.response.body.decode("unicode_escape")
		    ##print phone
		    #print 'Phone-OK'
		    #del g2
               #except (IndexError,GrabConnectionError,GrabNetworkError,GrabTimeoutError,ValueError,TypeError,AttributeError):
	            #phone = ''





               projects = {'sub': self.sub,
	                  'adress': mesto,
	                   'tip':tip,
	                   'naz':naz.replace(u'Сдам в аренду','').replace(u'Сдам','').replace(u'Аренда','').replace(u'Сдаётся','').replace(u'сдам','').replace(u'Сдаюся',''),
	                   'klass': klass,
	                   'cena': price,
	                   'plosh': plosh,
	                   'ohrana':ohrana,
	                   'gaz': gaz,
	                   'punkt':punkt,
	                   'voda': voda,
	                   'kanaliz': re.sub('[^\d\.]','',kanal),
	                   'electr': elek,
	                   'teplo': teplo,
	                   'opis': opis,
	                   'url': task.url,
	                   'phone': random.choice(list(open('../phone.txt').read().splitlines())),
	                   'lico':lico.replace(comp,''),
	                   'company': comp,
	                   'data':re.sub('[^\d\.]','',data)}

	       try:
		    link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+mesto
		    yield Task('adres',url=link,project=projects,refresh_cache=True,network_try_count=50)
	       except IndexError:
		    yield Task('adres',grab=grab,project=projects)



	  def task_adres(self, grab, task):

	       try:
		    ter=  grab.doc.rex_text(u'SubAdministrativeAreaName":"(.*?)"')
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

	       try:
		    lat=  grab.doc.rex_text(u'lowerCorner":"(.*?)"').split(' ')[0]
	       except IndexError:
	            lat =''
	       try:
	            lng=grab.doc.rex_text(u'lowerCorner":"(.*?)"').split(' ')[1]
	       except IndexError:
	            lng = ''

	       project2 ={'teritor': ter,
	                  'ulica':uliza,
	                  'dol':lat,
	                  'shir':lng,
	                  'dom':dom.replace('/','')}

	       yield Task('write',project=task.project,proj=project2,grab=grab)



	  def task_write(self,grab,task):
	       #time.sleep(1)
	       print('*'*100)
	       print  task.project['sub']
	       print  task.project['punkt']
	       print  task.proj['teritor']
	       print  task.proj['ulica']
	       print  task.proj['dom']
	       print  task.project['adress']
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






	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 24, task.project['adress'])
	       self.ws.write(self.result, 1, task.proj['teritor'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.proj['ulica'])
	       self.ws.write(self.result, 5, task.proj['dom'])
	       #self.ws.write(self.result, 9, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       #self.ws.write(self.result, 13, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 35, task.proj['dol'])
	       self.ws.write(self.result, 34, task.proj['shir'])
	       self.ws.write(self.result, 33, task.project['voda'])
	       self.ws.write(self.result, 30, task.project['kanaliz'])
	       #self.ws.write(self.result, 23, task.project['electr'])
	       #self.ws.write(self.result, 24, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'GDE.RU')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       #self.ws.write(self.result, 32, task.project['data1'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 28, oper)

	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',40,'***'
	       print oper
	       print('*'*100)
	       self.result+= 1





	       if str(self.result) == str(self.num):
		    self.stop()

     bot = move_Com(thread_number=10, network_try_limit=500)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=500, connect_timeout=500)
     try:
          bot.run()
     except KeyboardInterrupt:
          pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     bot.workbook.close()
     print('Done')


#time.sleep(5)
#os.system("/home/oleg/pars/gde/coma.py")


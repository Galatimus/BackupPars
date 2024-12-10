#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError
from grab import Grab
import logging
import os
import random
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



i = 86
l= open('links/com_arenda.txt').read().splitlines()
page = l[i]
oper = u'Аренда'




while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class move_Com(Spider):
	  def prepare(self):
	       self.f = page
	       self.link =l[i]

               while True:
                    try:
                         time.sleep(2)
			 g = Grab(timeout=20, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                         g.go(self.f)
			 if g.doc.select(u'//div[@class="total"]').exists()==True:
			      self.sub = g.doc.select(u'//div[@class="breadcrumbs"]/span[2]/a').text().replace('/','-')
			      self.num = re.sub('[^\d]','',g.doc.select(u'//div[@class="total"]/p').text())
			      self.pag = int(math.ceil(float(int(self.num))/float(20)))
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

	       self.workbook = xlsxwriter.Workbook(u'com/Move_%s' % bot.sub + u'_Коммерческая_'+oper+str(i)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'move_Коммерческая')
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
                    yield Task ('post',url=self.f+'?page=%d'%x+'&limit=20',refresh_cache=True,network_try_count=100)


	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@class="search-item__title-link search-item__item-link"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,refresh_cache=True, network_try_count=100)




	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title, " р-н")]').text()
	       except IndexError:
	            mesto =''

	       try:
		    try:
			 try:
			      try:
				   try:
					try:
					     try:
						  try:
						       punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"г.")]').text()
						  except IndexError:
						       punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"поселок")]').text()
					     except IndexError:
						  punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"д. ")]').text()
					except IndexError:
					     punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"сл.")]').text()
				   except IndexError:
					punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"пгт")]').text()
			      except IndexError:
				   punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"с.")]').text()
			 except IndexError:
			      punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"рп")]').text()
		    except IndexError:
			 punkt = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title,"п.")]').text()
	       except IndexError:
	            punkt =''



               try:
                    ter= grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "rayon")]').text()
               except IndexError:
                    ter =''


	       try:
		    try:
			 try:
			      try:
				   try:
		                        uliza = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "ulica")]').text()
		                   except IndexError:
			                uliza = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "prospekt")]').text()
		              except IndexError:
			           uliza = grab.doc.select(u'//div[@class="geo-block__geo-info_second-line"]/span[1]').text()
		         except IndexError:
			      uliza = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "proezd")]').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@href, "shosse")]').text()
	       except IndexError:
		    uliza =''


               try:
		    try:
                         dom = grab.doc.select(u'//a[@class="geo-block__geo-info_link"][contains(@title, "д.")]').text()
		    except IndexError:
		         dom = grab.doc.select(u'//div[@class="geo-block__geo-info_second-line"]/span[2]').text()
               except IndexError:
                    dom = ''

               try:
                    tip = grab.doc.select(u'//div[contains(text(),"Тип дома:")]/following-sibling::div').text()
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//div[contains(text(),"Тип объекта:")]/following-sibling::div').text()
               except IndexError:
                    naz =''
               try:
                    lin = []
		    for em in grab.doc.select(u'//li[@class="object-info__details-table_property"]/div[contains(@title, "г.")]'):
			 urr = em.text().replace(':','')
			 #print urr
			 lin.append(urr)
		    klass = ",".join(lin)
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//div[contains(text(),"Цена:")]/following-sibling::div').text()
               except IndexError:
                    price =''
               try:
                    plosh = grab.doc.select(u'//div[contains(text(),"Общая площадь:")]/following-sibling::div').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//div[contains(text()," Этаж:")]/following-sibling::div').text().split('/')[0]
               except IndexError:
                    ohrana =''
               try:
                    gaz =  grab.doc.select(u'//li[@class="geo-block__block-distance_property geo-block__block-distance_metro"]/a').text()
               except IndexError:
                    gaz =''
               try:
		    try:
                         voda =  grab.doc.select(u'//div[contains(text()," Этаж:")]/following-sibling::div').text().split('/')[1]
	            except IndexError:
		         voda =  grab.doc.select(u'//div[contains(text()," Количество этажей:")]/following-sibling::div').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//div[@class="object-place__address"]').text()#.split(u' на карте')[0].replace(u'Расположение ','')
               except IndexError:
                    kanal =''
               try:
                    elek = grab.doc.select(u'//div[contains(text(),"Год постройки:")]/following-sibling::div[1]').number()
               except IndexError:
                    elek =''

	       try:
	            zag = grab.doc.select(u'//h1').text()
	       except IndexError:
	            zag =''

               try:
                    lat = grab.doc.select(u'//script[@type="text/javascript"][contains(text(),"coordsCenterTile")]').text().split('coordsCenterTile=')[1].split('];')[0].replace('[','').split(',')[0]
               except IndexError:
	            lat =''

               try:
                    lng = grab.doc.select(u'//script[@type="text/javascript"][contains(text(),"coordsCenterTile")]').text().split('coordsCenterTile=')[1].split('];')[0].replace('[','').split(',')[1]
               except IndexError:
                    lng =''

	       try:
                    teplo = grab.doc.select(u'//li[@class="geo-block__block-distance_property geo-block__block-distance_walk-time"]').text()
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//h4[contains(text(),"Описание")]/following-sibling::div[1]').text()
	       except IndexError:
	            opis = ''
               try:
		    try:
		         lico = grab.doc.select(u'//div[@class="block-user__name"]').text()
		    except IndexError:
			 lico = grab.doc.select(u'//a[@class="block-user__name"]').text()
	       except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//div[@class="block-user__agency"]').text().replace(u'Собственник','')
               except IndexError:
                    comp = ''



               projects = {'sub': self.sub,
	                  'adress': mesto,
	                   'terit':ter,
	                   'punkt':punkt,
	                   'ulica':uliza,
	                   'dom':dom.replace('/','|'),
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
	                   'opis': re.sub('[=]','',opis),
	                   'url': task.url,
	                   'lico':lico,
	                   'company': comp,
	                   'shir':lat,
	                   'dol':lng,
	                   'zag':zag}


	       try:
		    link = task.url+'print/'
		    yield Task('phone',url=link,project=projects,refresh_cache=True,network_try_count=100)
	       except IndexError:
			 yield Task('phone',grab=grab,project=projects)


	  def task_phone(self, grab, task):
	       try:
		    phone= re.sub('[^\d\,]','',grab.doc.select(u'//div[@class="phone"]').text())
	       except IndexError:
		    phone = ''
	       try:
		    data1=  grab.doc.select(u'//div[@class="tech-info"]/div[2]/span').text().split(' ')[1]
	       except IndexError:
		    data1 =''
	       try:
		    data = grab.doc.select(u'//div[@class="tech-info"]/div[1]/span').text().split(' ')[1]
	       except IndexError:
		    data = ''

	       project2 ={'phone':phone,
	                  'dataraz': data,
	                  'dataob':data1}

	       yield Task('write',project=task.project,proj=project2,grab=grab)



	  def task_write(self,grab,task):
	       #if task.project['phone']<>'':
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
	       print  task.project['electr']
	       print  task.project['teplo']
	       print  task.project['ohrana']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.proj['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.proj['dataraz']
	       print  task.proj['dataob']
	       print  task.project['kanaliz']
	       print  task.project['zag']
	       print  task.project['shir']
	       print  task.project['dol']





	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['adress'])
	       self.ws.write(self.result, 3, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 5, task.project['dom'])
	       self.ws.write(self.result, 8, task.project['tip'])
	       self.ws.write(self.result, 9, task.project['naz'])
	       self.ws.write(self.result, 6, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 15, task.project['ohrana'])
	       self.ws.write(self.result, 26, task.project['gaz'])
	       self.ws.write(self.result, 16, task.project['voda'])
	       self.ws.write(self.result, 24, task.project['kanaliz'])
	       self.ws.write(self.result, 17, task.project['electr'])
	       self.ws.write(self.result, 27, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'MOVE.RU')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write(self.result, 21, task.proj['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 29, task.proj['dataraz'])
	       self.ws.write(self.result, 30, task.proj['dataob'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 28, oper)
	       self.ws.write(self.result, 33, task.project['zag'])
	       self.ws.write(self.result, 34, task.project['shir'])
	       self.ws.write(self.result, 35, task.project['dol'])
	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+self.num
	       logger.debug('Tasks - %s' % self.task_queue.size())
	       print '***',i+1,'/',len(l),'***'
	       print oper
	       print('*'*100)
	       self.result+= 1


	       if str(self.result) == str(self.num):
		    self.stop()


	       #if self.result >= 15:
	            #self.stop()



     bot = move_Com(thread_number=7, network_try_limit=1000)
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
          break




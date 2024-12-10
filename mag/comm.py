#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError
from grab import Grab
import logging
import math
import re
import os
import time
from datetime import datetime,timedelta
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)






i = 0
l= open('Links/Comm.txt').read().splitlines()
page = l[i]





while True:
     print '********************************************',i+1,'/',len(l),'*******************************************'
     class Mag_Com(Spider):
	  def prepare(self):
	       self.f = page
               for p in range(1,21):
                    try:
                         time.sleep(2)
			 g = Grab(timeout=20, connect_timeout=20)
			 g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                         g.go(self.f)
			 if 'prodazha' in self.f:
			      self.oper = u'Продажа' 
			 elif 'arenda' in self.f:
			      self.oper = u'Аренда'
                         #self.num =  re.sub('[^\d\,]','',g.doc.select(u'//title').text()).split(',')[0]
			 self.num = re.sub('[^\d]','',g.doc.select(u'//meta[@property="og:title"]').attr('content'))
                         self.sub = g.doc.select(u'//span[contains(text(),"Регион:")]/following-sibling::span').text()
			 
			 if self.num == '':
			      self.num = 1000
			 else:
			      self.num = self.num
			 
                         print self.sub
			 print self.num
			 print self.oper
			 del g
                         break
                    except(GrabTimeoutError,GrabNetworkError,IndexError,GrabConnectionError):
			 print g.config['proxy'],'Change proxy'
			 g.change_proxy()
			 del g
			 continue
	       else:
		    self.sub = ''
		    self.num=1	       

	       self.workbook = xlsxwriter.Workbook(u'com/Realtymag_'+self.oper+'_'+str(i+1)+'.xlsx')
	       self.ws = self.workbook.add_worksheet(u'Realtymag_Коммерческая')
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
	       self.ws.write(0, 36, u"ТРАССА")
	       self.ws.write(0, 37, u"ПАРКОВКА")
	       self.ws.write(0, 38, u"ОХРАНА")
	       self.ws.write(0, 39, u"КОНДИЦИОНИРОВАНИЕ")
	       self.ws.write(0, 40, u"ИНТЕРНЕТ")
	       self.ws.write(0, 41, u"ТЕЛЕФОН (КОЛИЧЕСТВО ЛИНИЙ)")
	       self.ws.write(0, 42, u"УСЛУГИ")
	       self.ws.write(0, 43, u"НАЛИЧИЕ ОТДЕЛКИ ПОМЕЩЕНИЙ")


	       self.result= 1





	  def task_generator(self):
	       yield Task ('post',url=self.f,refresh_cache=True,network_try_count=100)


	  def task_post(self,grab,task):
	       for elem in grab.doc.select('//a[@class="offer__details-link"]'):
		    ur = grab.make_url_absolute(elem.attr('href'))
		    #print ur
		    yield Task('item', url=ur,network_try_count=100)
	       yield Task("page", grab=grab,refresh_cache=True,network_try_count=100)

	  def task_page(self,grab,task):
	       try:
		    pg = grab.doc.select(u'//li[@class="new-pager__navigation-next"]/a')
		    u = grab.make_url_absolute(pg.attr('href'))
		    yield Task ('post',url= u,refresh_cache=True,network_try_count=100)
	       except IndexError:
		    print('*'*100)
		    print '!!!','NO PAGE NEXT','!!!'
		    print('*'*100)

	  def task_item(self, grab, task):
	       try:
	            mesto = grab.doc.select(u'//div[@class="offer-detail__district"]/a[contains(text(),"район")]').text()
	       except IndexError:
	            mesto =''

	       try:
	            punkt = grab.doc.select(u'//div[@class="offer-detail__city"]/a').text()
	       except IndexError:
	            punkt = ''

               try:
                    ter =  grab.doc.select(u'//div[@class="offer-detail__sublocality"]/a[contains(text(),"район")]').text()
               except IndexError:
                    ter =''
               try:
		    try:
                         uliza = grab.doc.select(u'//div[@class="offer-detail__address"]/a').text()
		    except IndexError:
			 uliza = grab.doc.select(u'//div[@class="offer-detail__address"]').text()
               except IndexError:
                    uliza = ''
               try:
		    try:
                         dom = grab.doc.select(u'//div[contains(text(),"Назначение")]/following-sibling::div').text()
		    except IndexError:
			 dom = grab.doc.select(u'//div[contains(text(),"Тип")]/following-sibling::div').text()
               except IndexError:
                    dom = ''

               try:
		    try:
                         tip = grab.doc.select(u'//div[contains(text(),"Размещение")]/following-sibling::div').text()
	            except IndexError:
		         tip = grab.doc.select(u'//div[contains(text(),"Тип здания")]/following-sibling::div').text()
               except IndexError:
                    tip = ''
               try:
                    naz = grab.doc.select(u'//h1').text().split(', ')[0]#.replace(u'офиса',u'офис').replace(u'офисного',u'офис').replace(u'офисное',u'офис')
               except IndexError:
                    naz =''
               try:
                    klass =  grab.doc.select(u'//div[contains(text(),"Класс")]/following-sibling::div').text()
               except IndexError:
                    klass = ''
               try:
                    price = grab.doc.select(u'//div[@class="offer-detail__price-rur"]').text()
               except IndexError:
                    price =''
               try:
                    plosh = grab.doc.select(u'//div[contains(text(),"Площадь")]/following-sibling::div').text()
               except IndexError:
                    plosh=''
               try:
                    ohrana = grab.doc.select(u'//div[contains(text(),"Этаж")]/following-sibling::div[contains(text(),"из")]').text().split(u' из ')[0]
               except IndexError:
                    ohrana =''
               try:
		    try:
                         gaz =  grab.doc.select(u'//div[contains(text(),"Этажность")]/following-sibling::div').text()
		    except IndexError:
			 gaz =  grab.doc.select(u'//div[contains(text(),"Этаж")]/following-sibling::div[contains(text(),"из")]').text().split(u' из ')[1]
               except IndexError:
                    gaz =''
               try:
                    voda =  grab.doc.select(u'//div[contains(text(),"Год постройки")]/following-sibling::div').text()
               except IndexError:
                    voda =''
               try:
                    kanal = grab.doc.select(u'//a[@class="offer-detail__metro-link"]').text()
               except IndexError:
                    kanal =''
               try:
                    elek = grab.doc.select(u'//div[contains(text(),"Удаленность")]/following-sibling::div').text()
               except IndexError:
                    elek =''
               try:
                    teplo = grab.doc.select(u'//h1').text()
               except IndexError:
                    teplo =''
               #time.sleep(1)
	       try:
		    opis = grab.doc.select(u'//div[@class="offer-detail__section-item section_type_additional-info"]').text()
	       except IndexError:
	            opis = ''
               try:
		    lico = grab.doc.select(u'//div[contains(text(),"частное лицо")]/preceding-sibling::div[@class="offer-detail__contact-name"]').text()
               except IndexError:
                    lico = ''
               try:
                    comp = grab.doc.select(u'//div[contains(text(),"агентство")]/preceding-sibling::div[@class="offer-detail__contact-name"]').text()
               except IndexError:
                    comp = ''

	       try:
	            con = [ (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
		              (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]

		    t= grab.doc.select(u'//div[@class="offer-detail__refresh"]').text()
		    if t.find(u'назад')>=0:
			 dt = u'вчера'
		    else:
		         dt= t
		    data = reduce(lambda dt, r: dt.replace(r[0], r[1]), con, dt)
	       except IndexError:
		    data=''

	       try:
                    phone = re.sub('[^\d\+]','',grab.doc.select(u'//div[@class="offer-detail__contact-phone-wrapper"]').attr('data-phone'))
               except IndexError:
	            phone = ''


	       try:
		    ohr = grab.doc.select(u'//div[contains(text(),"Охрана")]/following-sibling::div').text()
	       except IndexError:
		    ohr =''


	       if grab.doc.select(u'//div[contains(text(),"кондиционирования")]').exists() == True:
		    cond = 'Есть'
	       else:
	            cond = ''

	       try:
		    inet = grab.doc.select(u'//div[contains(text(),"в интернет")]/following-sibling::div').text()
	       except IndexError:
		    inet =''
	       try:
		    lat = grab.doc.rex_text(u'initMap(.*?);').split(',')[1]
	       except IndexError:
		    lat =''

	       try:
		    lng =  grab.doc.rex_text(u'initMap(.*?);').split(',')[2].replace(')','')
	       except IndexError:
		    lng =''
	       try:
		    lini = grab.doc.select(u'//div[contains(text(),"Установлено телефонных линий")]/following-sibling::div').text()#.split(u' из ')[1]
	       except IndexError:
		    lini =''


	       try:
		    usl = grab.doc.select(u'//div[contains(text(),"Общепит в здании")]/following-sibling::div').text()
	       except IndexError:
		    usl = ''

	       try:
		    otd = grab.doc.select(u'//div[contains(text(),"Состояние")]/following-sibling::div').text()
	       except IndexError:
	            otd = ''

	       try:
		    park = grab.doc.select(u'//span[contains(text(),"Парковка")]/following::span[1]').text()
	       except IndexError:
		    park = ''

	       try:
		    lin = []
		    for em in grab.doc.select(u'//div[@class="offer-detail__location-block"]/div'):
	                 urr = em.text()
	                 lin.append(urr)
	            rasp = ",".join(lin).replace(u'на карте,','')
	       except IndexError:
	            rasp =''


	       clear = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", opis)
	       clear = re.sub(u"[.,\-\s]{3,}", " ", clear)


               projects = {'sub': self.sub,
	                  'adress': mesto,
	                   'terit':ter,
	                   'punkt':punkt,
	                   'ulica':uliza,
	                   'dom':dom,
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
	                   'opis': clear,
	                   'url': task.url,
	                   'phone': phone,
	                   'lico':lico,
	                   'company': comp,
	                   'ohra':ohr,
	                   'condi':cond,
	                   'internet':inet,
	                   'shir': lat,
	                   'dol': lng,
	                   'linii': lini,
	                   'mesto': rasp.replace(kanal,''),
	                   'uslugi':usl,
	                   'sos': otd,
	                   'parkov': park,
	                   'data':data.replace(u'только что',datetime.today().strftime('%d.%m.%Y'))}



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
	       print  task.project['klass']
	       print  task.project['cena']
	       print  task.project['plosh']
	       print  task.project['gaz']
	       print  task.project['voda']
	       print  task.project['kanaliz']
	       print  task.project['electr']
	       print  task.project['ohrana']
	       print  task.project['opis']
	       print  task.project['url']
	       print  task.project['phone']
	       print  task.project['lico']
	       print  task.project['company']
	       print  task.project['data']
	       print  task.project['teplo']
	       print  task.project['ohra']
	       print  task.project['condi']
	       print  task.project['internet']
	       print  task.project['shir']
	       print  task.project['dol']
	       print  task.project['linii']
	       print  task.project['uslugi']
	       print  task.project['sos']
	       print  task.project['parkov']
	       print  task.project['mesto']





	       self.ws.write(self.result, 0, task.project['sub'])
	       self.ws.write(self.result, 1, task.project['adress'])
	       self.ws.write(self.result, 3, task.project['terit'])
	       self.ws.write(self.result, 2, task.project['punkt'])
	       self.ws.write(self.result, 4, task.project['ulica'])
	       self.ws.write(self.result, 9, task.project['dom'])
	       self.ws.write(self.result, 8, task.project['tip'])
	       self.ws.write(self.result, 7, task.project['naz'])
	       self.ws.write(self.result, 10, task.project['klass'])
	       self.ws.write(self.result, 11, task.project['cena'])
	       self.ws.write(self.result, 14, task.project['plosh'])
	       self.ws.write(self.result, 15, task.project['ohrana'])
	       self.ws.write(self.result, 16, task.project['gaz'])
	       self.ws.write(self.result, 17, task.project['voda'])
	       self.ws.write(self.result, 26, task.project['kanaliz'])
	       self.ws.write(self.result, 27, task.project['electr'])
	       self.ws.write(self.result, 33, task.project['teplo'])
	       self.ws.write(self.result, 18, task.project['opis'])
	       self.ws.write(self.result, 19, u'RealtyMag.RU')
	       self.ws.write_string(self.result, 20, task.project['url'])
	       self.ws.write_string(self.result, 21, task.project['phone'])
	       self.ws.write(self.result, 22, task.project['lico'])
	       self.ws.write(self.result, 23, task.project['company'])
	       self.ws.write(self.result, 29, task.project['data'])
	       self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	       self.ws.write(self.result, 28, self.oper)
	       self.ws.write(self.result, 38, task.project['ohra'])
	       self.ws.write(self.result, 39, task.project['condi'])
	       self.ws.write(self.result, 40, task.project['internet'])
	       self.ws.write(self.result, 34, task.project['shir'])
	       self.ws.write(self.result, 35, task.project['dol'])
	       self.ws.write(self.result, 41, task.project['linii'])
	       self.ws.write(self.result, 42, task.project['uslugi'])
	       self.ws.write(self.result, 43, task.project['sos'])
	       self.ws.write(self.result, 37, task.project['parkov'])
	       self.ws.write(self.result, 24, task.project['mesto'])



	       print('*'*100)
	       #print self.sub
	       print 'Ready - '+str(self.result)+'/'+str(self.num)
	       print 'Tasks - %s' % self.task_queue.size()
	       print '***',i+1,'/',len(l),'***'
	       print self.oper
	       print('*'*100)
	       self.result+= 1



	       if str(self.result) == str(self.num):
		    self.stop()


     bot = Mag_Com(thread_number=5, network_try_limit=1000)
     bot.load_proxylist('../tipa.txt','text_file')
     bot.create_grab_instance(timeout=50, connect_timeout=50)
     try:
          bot.run()
     except KeyboardInterrupt:
	  pass
     print('Wait 2 sec...')
     time.sleep(1)
     print('Save it...')
     bot.workbook.close()
     print('Done!')
     bot.shutdown()
     i=i+1
     try:
          page = l[i]
     except IndexError:
          break




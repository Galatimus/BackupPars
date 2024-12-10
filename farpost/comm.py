#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
from grab import Grab
import random
import re
from sub import conv
import time
import xlsxwriter
from datetime import datetime,timedelta
import sys
import os
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)



workbook = xlsxwriter.Workbook(u'0001-0070_00_C_001-0002_FARPOS.xlsx')


class Farpost_Com(Spider):
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
        #self.r = conv     

        self.result= 1



    def task_generator(self):
        l= open('faprost_com.txt').read().splitlines()
        self.dc = len(l)
        print self.dc
	headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
		  'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
		  'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0'}
	it = Grab(timeout=50, connect_timeout=50)
        for line in l:
	    #it.setup(url=line,headers=headers)
            yield Task ('item',url=line,refresh_cache=True,network_try_count=100)
	    #yield Task('item', grab=it,refresh_cache=True, network_try_count=100)
        

    def task_item(self, grab, task):

        try:
            ray = grab.doc.select(u'//div[@class="popButton cityPop"][contains(text(),"район")]').text()
        except IndexError:
            ray = ''          
        try:
            punkt= grab.doc.select(u'//div[@class="popButton cityPop"]').text()
            if 'район' in punkt:
                punkt = u'Владивосток'
            else:
                punkt = punkt
        except IndexError:
            punkt = ''

        try:
            ter= grab.doc.select(u'//div[contains(text(),"Район")]/following-sibling::div/span').text()
        except IndexError:
            ter =''

        try:
            uliza = grab.doc.select(u'//div[contains(text(),"Адрес")]/following-sibling::div/span/a[1]').text()
        except IndexError:
            uliza = ''

        try:
            dom = grab.doc.select(u'//title').text()
        except (IndexError,AttributeError):
            dom = ''

        try:
            trassa = grab.doc.select(u'//div[@id="breadcrumbs"]/div/span[5]/a').text().replace(u'Продажа ','').replace(u'Аренда ','')#.split(', ')[0]
        except IndexError:
            trassa = ''

        try:
            udal = grab.doc.select(u'//span[@itemprop="name"][contains(text(),"помещения")]').text()#.split(', ')[1]
        except IndexError:
            udal = ''
        try:
            seg = grab.doc.select(u'//span[@class="inplace"][contains(@data-field,"flatType")]').text()#.split(', ')[1]
        except IndexError:
            seg = ''	       

        try:
            try:
                price = grab.doc.select(u'//span[@data-field="price"]').text()#.replace(u'a',u'р.')
            except IndexError:
                price = grab.doc.select(u'//div[@class="viewbull-summary-price__realty-price"]').text()
        except IndexError:
            price = ''

        try:
            plosh = grab.doc.select(u'//span[@class="inplace"][contains(@data-field,"areaTotal")]').text()
        except IndexError:
            plosh = '' 
        try:
            cena_za = grab.doc.select(u'//span[@class="inplace"][contains(@data-field,"priceFor")]').text().replace(u'квадратный метр',u'м2').replace(u'все помещение','')
        except DataNotFound:
            cena_za = '' 


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
            teplo = grab.doc.select(u'//div[contains(text(),"Адрес")]/following-sibling::div/span').text().replace(u'Подробности о доме','')
        except DataNotFound:
            teplo =''  

        try:
            opis = grab.doc.select(u'//div[@class="bulletinText viewbull-field__container"]/p').text() 
        except IndexError:
            opis = ''



        try:
            try:
                oper = grab.doc.select(u'//div[@id="breadcrumbs"]/div/a[3]/span').text().split(' ')[0]
            except IndexError:
                oper = grab.doc.select(u'//div[@id="breadcrumbs"]/div/a[contains(@href,"garage")]').text().split(' ')[0]
        except IndexError:
            oper = ''

        try:

            con = [ ('сегодня', (datetime.today().strftime('%d.%m.%Y'))),
                    ('вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1))),
                    (' июля', '.07.2019'),(' июня', '.06.2019'),(' сентября', '.09.2018'),(' октября', '.10.2018'),(' января', '.01.2019'),(' февраля', '.02.2019'),(' марта', '.03.2019'),
                    (' мая', '.05.2019'),(' августа', '.08.2019'),(' декабря', '.12.2018'),(' апреля', '.04.2019'),(' ноября', '.11.2018')] 
            dt1= grab.doc.select(u'//span[@class="viewbull-header__actuality"]').text()
            data = reduce(lambda dt1, r1: dt1.replace(r1[0], r1[1]), con, dt1).replace(' ','')#.replace(u'более3-хмесяце', u'07.2015')
            #print data
        except IndexError:
            data = ''


        sub = reduce(lambda punkt, r: punkt.replace(r[0], r[1]), conv, punkt)

        projects = {'url': task.url,
                    'rayon': ray,
                    'sub': sub.replace(u' край край',' край'),
                    'punkt': punkt,
                    'teritor': ter,
                    'ulica': uliza.replace(u'Подробности о доме',''),
                    'dom': dom,
                    'trassa': trassa.replace(u'помещений',u'Помещение').replace(u'гаражей',u'Гараж'),
                    'udal': udal,
                    'segment': seg,
                    'cena': price,
                    'plosh':plosh,
                    #'phone':random.choice(list(open('../phone.txt').read().splitlines())),
                    'cena_za': cena_za.replace(u' в ',u'/'),
                    'ohrana':ohrana,
                    'gaz': gaz,
                    'voda': voda,
                    'kanaliz': kanal,
                    'electr': elek,
                    'teplo': teplo,
                    'opis':opis,
                    'operazia':oper,
                    'data':data}

        #yield Task('write',project=projects,grab=grab)
        #try:
            ##ad= grab.doc.select(u'//div[@class="popButton cityPop"]').text()+','+grab.doc.select(u'//div[contains(text(),"Адрес")]/following-sibling::div/span/a[1]').text()
            #ad= punkt+','+uliza
            #link = 'https://geocode-maps.yandex.ru/1.x/?format=json&geocode='+ad
            #yield Task('sub',url=link,project=projects,refresh_cache=True,network_try_count=100)
        #except IndexError:
            #yield Task('sub',grab=grab,project=projects)	  

        #try:
            ##ob = re.sub('[^\d]','',grab.doc.rex_text(u'>№(.+?)</div>'))
            #url_ph= grab.make_url_absolute(grab.doc.select(u'//div[@class="viewAjaxContactsPlaceHolder"]/a').attr('href'))
            #yield Task('phone',url=url_ph,project=projects,refresh_cache=True,network_try_count=100)
        #except IndexError:
        yield Task('phone',grab=grab,project=projects)

    def task_phone(self, grab, task):
        #try:
            #phone = re.sub('[^\d]','',grab.doc.select(u'//span[@class="phone"]').text())
            #lico = grab.doc.select(u'//span[@class="phone"]/following-sibling::text()').text()
        #except IndexError:
        phone = random.choice(list(open('../phone.txt').read().splitlines()))
        lico =''

        #yield Task('write',phone=phone,lico=lico,grab=grab)
	yield Task('write',project=task.project,phone=phone,lico=lico,grab=grab)

    def task_write(self,grab,task):
	if task.project['sub'] <> '':	      
	    print('*'*50)
	    print  task.project['sub']
	    print  task.project['rayon']
	    print  task.project['punkt']
	    print  task.project['teritor']
	    print  task.project['ulica']
	    print  task.project['dom']
	    print  task.project['trassa']
	    print  task.project['udal']
	    print  task.project['segment']
	    print  task.project['cena']+task.project['cena_za']
	    print  task.project['plosh']
	    print  task.project['ohrana']
	    print  task.project['gaz']
	    print  task.project['voda']
	    print  task.project['kanaliz']
	    print  task.project['electr']
	    print  task.project['opis']
	    print task.project['url']
	    print  task.phone
	    print  task.lico
	    print  task.project['data']
	    print  task.project['teplo']
    
	    #global result
	    self.ws.write(self.result, 0, task.project['sub'])
	    #self.ws.write(self.result, 1, task.project['rayon'])
	    self.ws.write(self.result, 2, task.project['punkt'])
	    self.ws.write(self.result, 3, task.project['teritor'])
	    self.ws.write(self.result, 4, task.project['ulica'])
	    self.ws.write(self.result, 7, task.project['segment'])
	    self.ws.write(self.result, 8, task.project['trassa'])
	    self.ws.write(self.result, 9, task.project['udal'])
	    self.ws.write(self.result, 33 , task.project['dom'])
	    self.ws.write(self.result, 11, task.project['cena']+task.project['cena_za'])
	    self.ws.write(self.result, 14, task.project['plosh'])
	    self.ws.write(self.result, 24, task.project['sub']+' ,'+task.project['teplo'])
	    self.ws.write(self.result, 19, u'FARPOST.RU')
	    self.ws.write_string(self.result, 20, task.project['url'])
	    self.ws.write(self.result, 18, task.project['opis'])
	    self.ws.write(self.result, 21, task.phone)
	    self.ws.write(self.result, 22, task.lico)
	    self.ws.write(self.result, 25, task.project['sub']+' ,'+task.project['teritor']+' ,'+task.project['teplo'])
	    self.ws.write(self.result, 29, task.project['data'])
	    self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	    self.ws.write(self.result, 28, task.project['operazia'])
	    print('*'*50)
	    #print task.sub
    
	    print 'Ready - '+str(self.result)+'/'+str(self.dc)
	    logger.debug('Tasks - %s' % self.task_queue.size())
	    #print '*',i+1,'/',dc,'*'
	    print  task.project['operazia']
	    print('*'*50)	       
	    self.result+= 1
    
    
    
	    #if self.result > 10:
		#self.stop()


bot = Farpost_Com(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5000, connect_timeout=5000)
try:
    bot.run()
except KeyboardInterrupt:
    pass
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
workbook.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/farpost/urlzem.py")









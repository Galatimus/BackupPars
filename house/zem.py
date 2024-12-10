#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import time
import re
import os
from datetime import datetime
import xlsxwriter
from grab import Grab
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)






i = 0
l= ['http://kalugahouse.ru/catalog/lands/prodazha/',
    'http://tulahouse.ru/catalog/lands/prodazha/',
    'http://vladimirhouse.ru/catalog/lands/prodazha/',
    'http://mohouse.ru/catalog/lands/prodazha/',
    'http://ryazanhouse.ru/catalog/lands/prodazha/']
     
page = l[i]  
oper = u'Продажа'



while True:
    print '********************************************',i+1,'/',len(l),'*******************************************'


    class House_Zem(Spider):



        def prepare(self):
            while True:
                try:
                    time.sleep(1)
                    g = Grab(timeout=20, connect_timeout=20) 
                    g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
                    g.go(page)
                    if 'kalugahouse' in page:
                        self.end = g.doc.select(u'//div[@id="right"]/preceding::a[1]').number()
                    elif 'tulahouse' in page:
                        self.end = g.doc.select(u'//div[@id="right"]/preceding::a[1]').number()
                    else:
                        self.end = g.doc.select(u'//td[@id="right"]/preceding::a[2]').number()
                    #try:
                        #self.end = g.doc.select(u'//td[@id="right"]/preceding::a[2]').number()
                    #except IndexError:
                        #self.end = g.doc.select(u'//div[@id="right"]/preceding::a[1]').number()
                    del g
                    break
                except(GrabTimeoutError,GrabNetworkError,GrabConnectionError):
                    print g.config['proxy'],'Change proxy'
                    g.change_proxy()
                    del g
                    continue
                except DataNotFound:
                    self.end = 1
                    del g
                    break
            
            conv = [ (u'ryazanhouse',u'Рязанская область'),(u'kalugahouse',u'Калужская область'),(u'tulahouse',u'Тульская область'),
                   (u'vladimirhouse',u'Владимировская область'),(u'mohouse',u'Московская область')]
            dt= re.findall('http://(.*?).ru',page)[0] 
            self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
            print self.sub,self.end
            self.workbook = xlsxwriter.Workbook(u'zem/House_%s' % bot.sub + u'_Земля_'+oper+'.xlsx')
            self.ws = self.workbook.add_worksheet(u'House_ЗЕМЛЯ')
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
            self.ws.write(0, 29, u"ДАТА_ПАРСИНГА")            
            self.result= 1
            self.conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
                    (u' мая ',u'.05.'),(u' июня ',u'.06.'),
                    (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
                    (u' января ',u'.01.'),(u' декабря ',u'.12.'),
                    (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
                    (u' февраля ',u'.02.'),(u' октября ',u'.10.')]            






        def task_generator(self):
            for x in range(1,self.end+1):
                yield Task ('post',url=page +'page_%d'% x,network_try_count=100)

        def task_post(self,grab,task):
            for elem in grab.doc.select(u'//div[@class="item-info"]/h2/a'):
                ur = grab.make_url_absolute(elem.attr('href'))  
                #print ur
                yield Task('item', url=ur,post = task.url,network_try_count=100)
                





        def task_item(self, grab, task):
            #pass

            try:
                ray = grab.doc.select(u'//td[@class="title"][contains(text(),"Район:")]/following-sibling::td').text()

            except DataNotFound:
                ray = ''
            try:
                punkt = grab.doc.select(u'//td[@class="title"][contains(text(),"Населенный пункт:")]/following-sibling::td').text().replace(self.sub,'')
            except IndexError:
                punkt = ''
            
            try:
                oren = grab.doc.select(u'//td[@class="title"][contains(text(),"Направление:")]/following-sibling::td').text()
                
            except IndexError:
                oren = ''
            try:
                udal = grab.doc.select(u'//td[@class="title"][contains(text(),"Удалённость:")]/following-sibling::td').text()
                
            except IndexError:
                udal = ''
            try:
                try:
                    price = grab.doc.select('//td[@class="title-price"]/following-sibling::td').text()
                except IndexError:
                    price = grab.doc.select('//td[@class="title-price"]/div').text()
            except IndexError:
                price = ''
            
            try:
                plosh = grab.doc.select(u'//td[@class="title"][contains(text(),"Площадь")]/following-sibling::td').text()
            except IndexError:
                plosh = ''
            try:
                vid = grab.doc.select(u'//ul[@id="land_data"]/li[contains(text(),"назначение")]').text().replace(u' назначение','')
            except IndexError:
                vid = ''
            try:
                gaz = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"газ")]').text().replace(u'есть газ',u'есть').replace(u'нет газа','')
            except IndexError:
                gaz =''
            try:
                voda = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"водоснабжение")]').text().replace(u'есть водоснабжение',u'есть').replace(u'нет водоснабжения','')
            except IndexError:
                voda =''
            try:
                kanal = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"канализация")]').text().replace(u'есть канализация',u'есть').replace(u'нет канализации','')
            except IndexError:
                kanal =''
            try:
                elekt = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"электричество")]').text().replace(u'есть электричество',u'есть').replace(u'нет электричества','')
            except IndexError:
                elekt =''
            try:
                teplo = re.sub(u'^.*(?=топление)','', grab.doc.select(u'//*[contains(text(), "топление")]').text())[:5].replace(u'топле',u'есть')
            except IndexError:
                teplo =''
            try:
                ohrana =re.sub(u'^.*(?=храна)','', grab.doc.select(u'//*[contains(text(), "храна")]').text())[:5].replace(u'храна',u'есть')
            except IndexError:
                ohrana =''
            try:
                opis = grab.doc.select(u'//p[@class="item-description"]').text() 
            except IndexError:
                opis = ''
            try:
                try:
                    ph = grab.doc.select(u'//td[@class="title"][contains(text(),"Телефон")]/following-sibling::td').text()
                    phone = re.sub('[^\d]', u'',ph)[:11]
                except IndexError:
                    ph = grab.doc.select(u'//td[@class="properties-phone"]').text()
                    phone = re.sub('[^\d]', u'',ph)[:11]
            except IndexError:
                phone = ''

            try:
                lico = grab.doc.select(u'//td[@class="title"][contains(text(),"Контактное лицо:")]/following-sibling::td').text()
            except IndexError:
                lico = ''

            try:
                com = grab.doc.select(u'//td[@class="title"][contains(text(),"Компания:")]/following-sibling::td').text()
            except IndexError:
                com = ''
            try:
                try:
                    d = grab.doc.select(u'//td[@class="title"][contains(text(),"Дата добавления:")]/following-sibling::td').text()
                    data = reduce(lambda d, r: d.replace(r[0], r[1]), self.conv, d).replace(u'г.','')
                except IndexError:
                    d = grab.doc.select(u'//p[contains(text(),"Размещено:")]').text().replace(u'Размещено: ','')
                    data = reduce(lambda d, r: d.replace(r[0], r[1]), self.conv, d).replace(u'г.','')                    
            except IndexError:
                data = ''
            try:
                conv1 = [ (u'ryazanhouse',u'Ryazanhouse.ru'),(u'kalugahouse',u'Kalugahouse.ru'),
                             (u'tulahouse',u'Tulahouse.ru'),(u'vladimirhouse',u'Vladimirhouse.ru'),
                             (u'mohouse',u'Mohouse.ru')]
                dt= re.findall('http://(.*?).ru',task.url)[0] 
                istoch = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv1, dt)
            except IndexError:
                istoch = ''                

            

            projects = {'url': task.url,
                        'sub': self.sub,
                        'rayon': ray,
                        'punkt': punkt,
                        'oren': oren,
                        'udal': udal,
                        'price': price,
                        'ploshad': plosh,
                        'vid': vid,
                        'gaz': gaz,
                        'voda':voda,
                        'kanal': kanal,
                        'elekt': elekt,
                        'teplo': teplo,
                        'ohrana': ohrana,
                        'opis': opis,
                        'phone': phone,
                        'lico':lico,
                        'company':com,
                        'dataraz': data,
                        'istochnik':istoch
                        
                        }



            yield Task('write',project=projects,grab=grab)






        def task_write(self,grab,task):

            print('*'*100)
            print  task.project['sub']
            print  task.project['rayon']
            print  task.project['punkt']
            print  task.project['oren']
            print  task.project['udal']
            print  task.project['price']
            print  task.project['ploshad']
            print  task.project['vid']
            print  task.project['gaz']
            print  task.project['voda']
            print  task.project['kanal']
            print  task.project['elekt']
            print  task.project['teplo']
            print  task.project['ohrana']
            print  task.project['opis']
            print task.project['url']
            print  task.project['phone']
            print  task.project['lico']
            print  task.project['company']
            print  task.project['dataraz']
            print  task.project['istochnik']
            

            self.ws.write(self.result, 0, task.project['sub'])
            self.ws.write(self.result, 1, task.project['rayon'])
            self.ws.write(self.result, 2, task.project['punkt'])
            #self.ws.write(self.result, 4, task.project['ulica'])
            self.ws.write(self.result, 6, task.project['oren'])
            self.ws.write(self.result, 8, task.project['udal'])
            self.ws.write(self.result, 9, oper)
            self.ws.write(self.result, 10, task.project['price'])
            #self.ws.write(self.result, 11, task.project['price_sot'])
            self.ws.write(self.result, 12, task.project['ploshad'])
            self.ws.write(self.result, 14, task.project['vid'])
            self.ws.write(self.result, 15, task.project['gaz'])
            self.ws.write(self.result, 16, task.project['voda'])
            self.ws.write(self.result, 17, task.project['kanal'])
            self.ws.write(self.result, 18, task.project['elekt'])
            self.ws.write(self.result, 19, task.project['teplo'])
            self.ws.write(self.result, 20, task.project['ohrana'])
            self.ws.write(self.result, 22, task.project['opis'])
            self.ws.write(self.result, 23, task.project['istochnik'])
            self.ws.write_string(self.result, 24, task.project['url'])
            self.ws.write(self.result, 25, task.project['phone'])
            self.ws.write(self.result, 26, task.project['lico'])
            self.ws.write(self.result, 27, task.project['company'])
            self.ws.write(self.result, 28, task.project['dataraz'])
            self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
            




            print('*'*100)
            print self.sub
            print 'Ready - '+str(self.result)
            logger.debug('Tasks - %s' % self.task_queue.size()) 
            print '***',i+1,'/',len(l),'***'
            print oper
            print('*'*100)
            self.result+= 1

            #if self.result > 10:
                #self.stop()

    
    bot = House_Zem(thread_number=5, network_try_limit=1000)
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
        if oper == u'Продажа':
            i = 0
            l=['http://ryazanhouse.ru/catalog/lands/arenda/',
               'http://kalugahouse.ru/catalog/lands/arenda/',
               'http://tulahouse.ru/catalog/lands/arenda/',
               'http://vladimirhouse.ru/catalog/lands/arenda/',
               'http://mohouse.ru/catalog/lands/arenda/']
            dc = len(l)
            page = l[i]  
            oper = u'Аренда'
        else:
            break
time.sleep(5)
os.system("/home/oleg/pars/house/com.py")
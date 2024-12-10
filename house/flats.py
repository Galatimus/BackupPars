#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import time
import re
from datetime import datetime
import xlsxwriter
from grab import Grab
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


g = Grab(timeout=20, connect_timeout=20) 



i = 0
l= ['http://ryazanhouse.ru/catalog/kvartiry/prodazha/',
    'http://kalugahouse.ru/catalog/appartments/prodazha/',
    'http://tulahouse.ru/catalog/appartments/prodazha/',
    'http://vladimirhouse.ru/catalog/appartments/prodazha/',
    'http://mohouse.ru/catalog/kvartiry/prodazha/',
    'http://ryazanhouse.ru/catalog/kvartiry/prodazha_komnat/',
    'http://kalugahouse.ru/catalog/appartments/prodazha_komnat/',
    'http://tulahouse.ru/catalog/appartments/prodazha_komnat/',
    'http://vladimirhouse.ru/catalog/appartments/prodazha_komnat/',
    'http://mohouse.ru/catalog/kvartiry/prodazha_komnat/']
     
page = l[i]  
oper = u'Продажа'



while True:
    print '********************************************',i+1,'/',len(l),'*******************************************'


    class House_KV(Spider):



        def prepare(self):
            while True:
                try:
                    time.sleep(1)
                    g.go(page)
                    try:
                        self.end = g.doc.select(u'//td[@id="right"]/preceding::a[2]').number()
                    except IndexError:
                        self.end = g.doc.select(u'//div[@id="right"]/preceding::a[2]').number()
                    break
                except(GrabTimeoutError,GrabNetworkError,GrabConnectionError):
                    print g.config['proxy'],'Change proxy'
                    g.change_proxy()
                    continue
                except DataNotFound:
                    self.end = 1
                    break
            
            conv = [ (u'ryazanhouse',u'Рязанская область'),(u'kalugahouse',u'Калужская область'),(u'tulahouse',u'Тульская область'),
                   (u'vladimirhouse',u'Владимировская область'),(u'mohouse',u'Московская область')]
            dt= re.findall('http://(.*?).ru',page)[0] 
            self.sub = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
            print self.sub,self.end
            self.workbook = xlsxwriter.Workbook(u'Kv/House_%s' % bot.sub + u'_Жилье_'+oper+str(i+1) +'.xlsx')
            self.ws = self.workbook.add_worksheet(u'House_Жилье')
            self.ws.write(0, 0, "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
            self.ws.write(0, 1, "МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
            self.ws.write(0, 2, "НАСЕЛЕННЫЙ_ПУНКТ")
            self.ws.write(0, 3, "ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
            self.ws.write(0, 4, "УЛИЦА")
            self.ws.write(0, 5, "ДОМ")
            self.ws.write(0, 6, "ОРИЕНТИР")
            self.ws.write(0, 7, "СТАНЦИЯ_МЕТРО")
            self.ws.write(0, 8, "ДО_МЕТРО_МИНУТ")
            self.ws.write(0, 9, "ПЕШКОМ_ТРАНСПОРТОМ")
            self.ws.write(0, 10, "ТИП_ОБЪЕКТА")
            self.ws.write(0, 11, "ОПЕРАЦИЯ")
            self.ws.write(0, 12, "СТОИМОСТЬ")
            self.ws.write(0, 13, "ЦЕНА_М2")
            self.ws.write(0, 14, "КОЛИЧЕСТВО_КОМНАТ")
            self.ws.write(0, 15, "ПЛОЩАДЬ_ОБЩАЯ")
            self.ws.write(0, 16, "ПЛОЩАДЬ_ЖИЛАЯ")
            self.ws.write(0, 17, "ПЛОЩАДЬ_КУХНИ")
            self.ws.write(0, 18, "ПЛОЩАДЬ_КОМНАТ")
            self.ws.write(0, 19, "ЭТАЖ")
            self.ws.write(0, 20, "ЭТАЖНОСТЬ")
            self.ws.write(0, 21, "МАТЕРИАЛ_СТЕН")
            self.ws.write(0, 22, "ГОД_ПОСТРОЙКИ")
            self.ws.write(0, 23, "РАСПОЛОЖЕНИЕ_КОМНАТ")
            self.ws.write(0, 24, "БАЛКОН")
            self.ws.write(0, 25, "ЛОДЖИЯ")
            self.ws.write(0, 26, "САНУЗЕЛ")
            self.ws.write(0, 27, "ОКНА")
            self.ws.write(0, 28, "СОСТОЯНИЕ")
            self.ws.write(0, 29, "ВЫСОТА_ПОТОЛКОВ")
            self.ws.write(0, 30, "ЛИФТ")
            self.ws.write(0, 31, "РЫНОК")
            self.ws.write(0, 32, "КОНСЬЕРЖ")
            self.ws.write(0, 33, "ОПИСАНИЕ")
            self.ws.write(0, 34, "ИСТОЧНИК_ИНФОРМАЦИИ")
            self.ws.write(0, 35, "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
            self.ws.write(0, 36, "ТЕЛЕФОН")
            self.ws.write(0, 37, "КОНТАКТНОЕ_ЛИЦО")
            self.ws.write(0, 38, "КОМПАНИЯ")
            self.ws.write(0, 39, "ДАТА_РАЗМЕЩЕНИЯ_ОБЪЯВЛЕНИЯ")
            self.ws.write(0, 40, "ДАТА_ПАРСИНГА")
            self.result= 1






        def task_generator(self):
            for x in range(1,self.end+1):
                yield Task ('post',url=page +'page_%d'%x,network_try_count=100)

        def task_post(self,grab,task):
            for elem in grab.doc.select(u'//div[@class="item-info"]/h2/a'):
                ur = grab.make_url_absolute(elem.attr('href'))  
                #print ur
                yield Task('item', url=ur,post = task.url,network_try_count=100)
                





        def task_item(self, grab, task):
            #pass

            try:
                ray = ''#grab.doc.select(u'//td[@class="title"][contains(text(),"Район:")]/following-sibling::td').text()

            except DataNotFound:
                ray = ''
            try:
                punkt = grab.doc.select(u'//td[@class="title"][contains(text(),"Населенный пункт:")]/following-sibling::td').text()
            except IndexError:
                punkt = ''
                
            try:
                try:
                    ter = grab.doc.select(u'//td[@class="title"][contains(text(),"Район:")]/following-sibling::td').text()
                except IndexError:
                    ter = grab.doc.select(u'//td[@class="title"][contains(text(),"Улица (деревня):")]/following-sibling::td').text()               
            except IndexError:
                ter=''                
            
            try:
                uliza = grab.doc.select(u'//td[@class="title"][contains(text(),"Улица:")]/following-sibling::td').text()
            except IndexError:
                uliza = ''
            try:
                dom = grab.doc.select(u'//td[@class="title"][contains(text(),"Номер дома:")]/following-sibling::td').text()
            except IndexError:
                dom = ''
            try:
                tip_ob=grab.doc.select(u'//div[@id="nav"]/a[4]').text().split(' ')[1].replace(u'квартир',u'Квартира').replace(u'комнат',u'Комната')
            except IndexError:
                tip_ob = ''                      
            try:
                try:
                    price = grab.doc.select('//td[@class="title-price"]/following-sibling::td').text()
                except IndexError:
                    price = grab.doc.select('//td[@class="title-price"]/div').text()
            except IndexError:
                price = ''
            try:
                kol_komnat = grab.doc.select(u'//td[@class="title"][contains(text(),"Комнат:")]/following-sibling::td').number()
            except IndexError:
                kol_komnat = ''                
            
            try:
                plosh = grab.doc.select(u'//td[@class="title"][contains(text(),"Общая площадь:")]/following-sibling::td').text()
            except IndexError:
                plosh = ''
            try:
                plosh_gil = grab.doc.select(u'//td[@class="title"][contains(text(),"Жилая площадь:")]/following-sibling::td').text()
            except IndexError:
                plosh_gil = '' 
            try:
                plosh_kuh = grab.doc.select(u'//td[@class="title"][contains(text(),"Кухня площадь:")]/following-sibling::td').text()
            except IndexError:
                plosh_kuh = ''
            try:
                et = grab.doc.select(u'//td[@class="title"][contains(text(),"Этаж:")]/following-sibling::td').number()
            except IndexError:
                et = '' 
            try:
                etagn = grab.doc.select(u'//td[@class="title"][contains(text(),"Всего этажей:")]/following-sibling::td').number()
            except IndexError:
                etagn = ''
            try:
                mat = grab.doc.select(u'//td[@class="title"][contains(text(),"Тип дома:")]/following-sibling::td').text()
            except IndexError:
                mat = ''                
            try:
                balkon = grab.doc.select(u'//td[@class="title"][contains(text(),"Балкон, лоджия:")]/following-sibling::td').text().replace(u'балкон',u'есть').replace(u'лоджия','').replace(u'Нет','')
            except IndexError:
                balkon = ''
            try:
                lodg = grab.doc.select(u'//td[@class="title"][contains(text(),"Балкон, лоджия:")]/following-sibling::td').text().replace(u'лоджия',u'есть').replace(u'балкон','').replace(u'Нет','')
            except IndexError:
                lodg = ''
            try:
                sost = grab.doc.select(u'//td[@class="title"][contains(text(),"Планировка:")]/following-sibling::td').text()
            except IndexError:
                sost = ''
            try:
                rinok = grab.doc.select(u'//td[@class="title"][contains(text(),"Тип:")]/following-sibling::td').text()
            except IndexError:
                rinok = ''
            try:
                kons = re.sub(u'^.*(?=консьерж)','', grab.doc.select(u'//*[contains(text(), "консьерж")]').text())[:8].replace(u'консьерж',u'есть')
            except IndexError:
                kons = ''
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
                conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
                        (u' мая ',u'.05.'),(u' июня ',u'.06.'),
                        (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
                        (u' января ',u'.01.'),(u' декабря ',u'.12.'),
                        (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
                        (u' февраля ',u'.02.'),(u' октября ',u'.10.')] 
                d = grab.doc.select(u'//td[@class="title"][contains(text(),"Дата добавления:")]/following-sibling::td').text()
                data = reduce(lambda d, r: d.replace(r[0], r[1]), conv, d).replace(u'г.','')
            except IndexError:
                data = ''
            try:
                conv = [ (u'ryazanhouse',u'Ryazanhouse.ru'),(u'kalugahouse',u'Kalugahouse.ru'),
                         (u'tulahouse',u'Tulahouse.ru'),(u'vladimirhouse',u'Vladimirhouse.ru'),
                         (u'mohouse',u'Mohouse.ru')]
                dt= re.findall('http://(.*?).ru',task.url)[0] 
                istoch = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)
            except IndexError:
                istoch = '' 


            projects = {'sub': self.sub,
                        'rayon': ray,
                        'punkt': punkt,
                        'teritor': ter,
                        'ulica': uliza,
                        'dom': dom,
                        'object': tip_ob,
                        'cena': price,
                        'col_komnat': kol_komnat,
                        'plosh_ob':plosh,
                        'plosh_gil': plosh_gil,
                        'plosh_kuh': plosh_kuh,
                        'etach': et,
                        'etashost': etagn,
                        'material': mat,
                        'balkon': balkon,
                        'logia': lodg,
                        'sost': sost,
                        'rinok': rinok,
                        'kons':kons,
                        'opis':opis,
                        'url':task.url,
                        'phone':phone,
                        'lico':lico,
                        'company':com,
                        'data':data,
                        'istochnik':istoch}



            yield Task('write',project=projects,grab=grab)






        def task_write(self,grab,task):

            print('*'*100)
            print  task.project['sub']
            print  task.project['rayon']
            print  task.project['punkt']
            print  task.project['teritor']
            print  task.project['ulica']
            print  task.project['dom']
            #print  task.project['metro']
            #print  task.project['udall']
            print  task.project['cena']
            print  task.project['col_komnat']
            print  task.project['plosh_ob']
            print  task.project['plosh_gil']
            print  task.project['plosh_kuh']
            print  task.project['etach']
            print  task.project['etashost']
            print  task.project['material']
            #print  task.project['god_post']
            print  task.project['balkon']
            print  task.project['logia']
            #print  task.project['uzel']
            print  task.project['sost']
            #print  task.project['vis_potolok']
            #print  task.project['lift']
            print  task.project['rinok']
            print  task.project['kons']
            print  task.project['opis']
            print  task.project['url']
            print  task.project['phone']
            print  task.project['lico']
            print  task.project['company']
            print  task.project['data']
            print  task.project['istochnik']
            print  task.project['object']

            self.ws.write(self.result, 0, task.project['sub'])
            self.ws.write(self.result, 1, task.project['rayon'])
            self.ws.write(self.result, 2, task.project['punkt'])
            self.ws.write(self.result, 3, task.project['teritor'])
            self.ws.write(self.result, 4, task.project['ulica'])
            self.ws.write(self.result, 5, task.project['dom'])
            #self.ws.write(self.result, 7, task.project['metro'])
            #self.ws.write(self.result, 9, task.project['udall'])
            self.ws.write(self.result, 10, task.project['object'])
            self.ws.write(self.result, 11, oper)
            self.ws.write(self.result, 12, task.project['cena'])
            self.ws.write(self.result, 14, task.project['col_komnat'])
            self.ws.write(self.result, 15, task.project['plosh_ob'])
            self.ws.write(self.result, 16, task.project['plosh_gil'])
            self.ws.write(self.result, 17, task.project['plosh_kuh'])
            self.ws.write(self.result, 19, task.project['etach'])
            self.ws.write(self.result, 20, task.project['etashost'])
            self.ws.write(self.result, 21, task.project['material'])
            #self.ws.write(self.result, 22, task.project['god_post'])
            self.ws.write(self.result, 24, task.project['balkon'])
            self.ws.write(self.result, 25, task.project['logia'])
            #self.ws.write(self.result, 26, task.project['uzel'])
            self.ws.write(self.result, 28, task.project['sost'])
            #self.ws.write(self.result, 29, task.project['vis_potolok'])
            #self.ws.write(self.result, 30, task.project['lift'])
            self.ws.write(self.result, 31, task.project['rinok'])
            self.ws.write(self.result, 32, task.project['kons'])
            self.ws.write(self.result, 33, task.project['opis'])
            self.ws.write(self.result, 34, task.project['istochnik'])
            self.ws.write_string(self.result, 35, task.project['url'])
            self.ws.write(self.result, 36, task.project['phone'])
            self.ws.write(self.result, 37, task.project['lico'])
            self.ws.write(self.result, 38, task.project['company'])
            self.ws.write(self.result, 39, task.project['data'])
            self.ws.write(self.result, 40, datetime.today().strftime('%d.%m.%Y'))
            




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

    
    bot = House_KV(thread_number=3, network_try_limit=2000)
    bot.load_proxylist('/home/oleg/Proxy/tipa.txt','text_file')
    bot.create_grab_instance(timeout=5000, connect_timeout=5000)
    bot.run()

    print bot.sub
    print(u'Спим 2 сек...')
    time.sleep(2)
    print(u'Сохранение...')
    bot.workbook.close()
    print('Done!')
   

    i=i+1
    try:
        page = l[i]
    except IndexError:
        if oper == u'Продажа':
            i = 0
            l= ['http://ryazanhouse.ru/catalog/kvartiry/arenda/',
                'http://kalugahouse.ru/catalog/appartments/arenda/',
                'http://tulahouse.ru/catalog/appartments/arenda/',
                'http://vladimirhouse.ru/catalog/appartments/arenda/',
                'http://mohouse.ru/catalog/kvartiry/arenda/',
                'http://ryazanhouse.ru/catalog/kvartiry/arenda_komnat/',
                'http://kalugahouse.ru/catalog/appartments/arenda_komnat/',
                'http://tulahouse.ru/catalog/appartments/arenda_komnat/',
                'http://vladimirhouse.ru/catalog/appartments/arenda_komnat/',
                'http://mohouse.ru/catalog/kvartiry/arenda_komnat/']
            dc = len(l)
            page = l[i]  
            oper = u'Аренда'
        else:
            break



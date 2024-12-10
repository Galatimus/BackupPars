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
l= ['http://ryazanhouse.ru/catalog/houses/prodazha/',
    'http://kalugahouse.ru/catalog/houses/prodazha/',
    'http://tulahouse.ru/catalog/houses/prodazha/',
    'http://vladimirhouse.ru/catalog/houses/prodazha/',
    'http://mohouse.ru/catalog/doma/prodazha/']
     
page = l[i]  
oper = u'Продажа'



while True:
    print '********************************************',i+1,'/',len(l),'*******************************************'


    class House_Zag(Spider):



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
            self.workbook = xlsxwriter.Workbook(u'zag/House_%s' % bot.sub + u'_Загород_'+oper+'.xlsx')
            self.ws = self.workbook.add_worksheet(u'House_Загород')
            self.ws.write(0, 0, "СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
            self.ws.write(0, 1, "МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
            self.ws.write(0, 2, "НАСЕЛЕННЫЙ_ПУНКТ")
            self.ws.write(0, 3, "ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
            self.ws.write(0, 4, "УЛИЦА")
            self.ws.write(0, 5, "ДОМ")
            self.ws.write(0, 6, "ОРИЕНТИР")
            self.ws.write(0, 7, "ТРАССА")
            self.ws.write(0, 8, "УДАЛЕННОСТЬ")
            self.ws.write(0, 9, "КАДАСТРОВЫЙ_НОМЕР_ЗЕМЕЛЬНОГО_УЧАСТКА")
            self.ws.write(0, 10, "ТИП_ОБЪЕКТА")
            self.ws.write(0, 11, "ОПЕРАЦИЯ")
            self.ws.write(0, 12, "СТОИМОСТЬ")
            self.ws.write(0, 13, "ЦЕНА_М2")
            self.ws.write(0, 14, "ПЛОЩАДЬ_ОБЩАЯ")
            self.ws.write(0, 15, "КОЛИЧЕСТВО_КОМНАТ")
            self.ws.write(0, 16, "ЭТАЖНОСТЬ")
            self.ws.write(0, 17, "МАТЕРИАЛ_СТЕН")
            self.ws.write(0, 18, "ГОД_ПОСТРОЙКИ")
            self.ws.write(0, 19, "ПЛОЩАДЬ_УЧАСТКА")
            self.ws.write(0, 20, "ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
            self.ws.write(0, 21, "ГАЗОСНАБЖЕНИЕ")
            self.ws.write(0, 22, "ВОДОСНАБЖЕНИЕ")
            self.ws.write(0, 23, "КАНАЛИЗАЦИЯ")
            self.ws.write(0, 24, "ЭЛЕКТРОСНАБЖЕНИЕ")
            self.ws.write(0, 25, "ТЕПЛОСНАБЖЕНИЕ")
            self.ws.write(0, 26, "ЛЕС")
            self.ws.write(0, 27, "ВОДОЕМ")
            self.ws.write(0, 28, "БЕЗОПАСНОСТЬ")
            self.ws.write(0, 29, "ОПИСАНИЕ")
            self.ws.write(0, 30, "ИСТОЧНИК_ИНФОРМАЦИИ")
            self.ws.write(0, 31, "ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
            self.ws.write(0, 32, "ТЕЛЕФОН")
            self.ws.write(0, 33, "КОНТАКТНОЕ_ЛИЦО")
            self.ws.write(0, 34, "КОМПАНИЯ")
            self.ws.write(0, 35, "ДАТА_РАЗМЕЩЕНИЯ")
            self.ws.write(0, 36, "ДАТА_ПАРСИНГА")
            self.result= 1






        def task_generator(self):
            for x in range(1,self.end+1):
                yield Task ('post',url=page +'page_%d'% x,network_try_count=100)

        def task_post(self,grab,task):
            for elem in grab.doc.select(u'//div[@class="item-info"]/h2/a'):
                ur = grab.make_url_absolute(elem.attr('href'))  
                #print ur
                yield Task('item', url=ur,post = task.url,network_try_count=100,use_proxylist=False)
                





        def task_item(self, grab, task):
            #pass

            try:
                ray = grab.doc.select(u'//td[@class="title"][contains(text(),"Район:")]/following-sibling::td').text()

            except IndexError:
                ray = ''
            try:
                punkt = grab.doc.select(u'//td[@class="title"][contains(text(),"Населенный пункт:")]/following-sibling::td').text()
            except IndexError:
                punkt = ''
            try:
                ter = grab.doc.select(u'//td[@class="title"][contains(text(),"Улица (деревня):")]/following-sibling::td').text()               
            except IndexError:
                ter=''
            try:
                oren = grab.doc.select(u'//td[@class="title"][contains(text(),"Направление:")]/following-sibling::td').text()
                
            except IndexError:
                oren = ''
            try:
                udal = grab.doc.select(u'//td[@class="title"][contains(text(),"Удалённость:")]/following-sibling::td').text()
                
            except IndexError:
                udal = ''
            try:
                t = grab.doc.select(u'//div[@class="item"]/h1').text()
                if t.find(u'оттедж')>=0:
                    tip_ob = u'Коттедж'
                elif t.find(u'Дача')>=0:
                    tip_ob = u'Дача'
                elif t.find(u'аунхаус')>=0:
                    tip_ob = u'Таунхаус'    
                else:
                    tip_ob = u'Дом'                
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
                plosh = grab.doc.select(u'//td[@class="title"][contains(text(),"Общая площадь:")]/following-sibling::td').text()
            except IndexError:
                plosh = ''
            try:
                etash = grab.doc.select(u'//td[@class="title"][contains(text(),"Этажей")]/following-sibling::td').number()
            except IndexError:
                etash = ''
            try:
                plosh_uch = grab.doc.select(u'//td[@class="title"][contains(text(),"Площадь участка:")]/following-sibling::td').text()
            except IndexError:
                plosh_uch = ''                
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

            

            projects = {'url': task.url,
                        'sub': self.sub,
                        'rayon': ray,
                        'punkt': punkt,
                        'teritor': ter,
                        'oren': oren,
                        'udal': udal,
                        'object': tip_ob,
                        'price': price,
                        'ploshad': plosh,
                        'plouh': plosh_uch,
                        'vid': etash,
                        'ohrana':ohrana,
                        'gaz': gaz,
                        'voda': voda,
                        'kanaliz': kanal,
                        'electr': elek,
                        'teplo': teplo,
	                'les': les,
                        'vodoem':vodoem,	 
                        'opis': opis,
                        'phone': phone,
                        'lico':lico,
                        'company':com,
                        'dataraz': data,
                        'istochnik':istoch}



            yield Task('write',project=projects,grab=grab)






        def task_write(self,grab,task):

            print('*'*100)
            print  task.project['sub']
            print  task.project['rayon']
            print  task.project['punkt']
            print  task.project['teritor']
            print  task.project['oren']
            print  task.project['udal']
            print  task.project['object']
            print  task.project['price']
            print  task.project['ploshad']
            print  task.project['plouh']
            print  task.project['vid']
            print  task.project['gaz']
            print  task.project['voda']
            print  task.project['kanaliz']
            print  task.project['electr']
            print  task.project['teplo']
            print  task.project['ohrana']
            print  task.project['les']
            print  task.project['vodoem']	             
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
            self.ws.write(self.result, 3, task.project['teritor'])
            self.ws.write(self.result, 6, task.project['oren'])
            self.ws.write(self.result, 8, task.project['udal'])
            self.ws.write(self.result, 11, oper)
            self.ws.write(self.result, 10, task.project['object'])
            self.ws.write(self.result, 12, task.project['price'])
            self.ws.write(self.result, 14, task.project['ploshad'])
            self.ws.write(self.result, 16, task.project['vid'])
            self.ws.write(self.result, 19, task.project['plouh'])
            self.ws.write(self.result, 21, task.project['gaz'])
            self.ws.write(self.result, 22, task.project['voda'])
            self.ws.write(self.result, 23, task.project['kanaliz'])
            self.ws.write(self.result, 24, task.project['electr'])
            self.ws.write(self.result, 25, task.project['teplo'])
            self.ws.write(self.result, 28, task.project['ohrana'])
            self.ws.write(self.result, 26, task.project['les'])
            self.ws.write(self.result, 27, task.project['vodoem'])            
            self.ws.write(self.result, 29, task.project['opis'])
            self.ws.write(self.result, 30, task.project['istochnik'])
            self.ws.write_string(self.result, 31, task.project['url'])
            self.ws.write(self.result, 32, task.project['phone'])
            self.ws.write(self.result, 33, task.project['lico'])
            self.ws.write(self.result, 34, task.project['company'])
            self.ws.write(self.result, 35, task.project['dataraz'])
            self.ws.write(self.result, 36, datetime.today().strftime('%d.%m.%Y'))
            




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

    
    bot = House_Zag(thread_number=3, network_try_limit=1000)
    bot.load_proxylist('/home/oleg/Proxy/tipa.txt','text_file')
    bot.create_grab_instance(timeout=100, connect_timeout=1000)
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
            l= ['http://ryazanhouse.ru/catalog/houses/arenda_predlozheniya/',
                'http://kalugahouse.ru/catalog/houses/arenda_predlozheniya/',
                'http://tulahouse.ru/catalog/houses/arenda_predlozheniya/',
                'http://vladimirhouse.ru/catalog/houses/arenda_predlozheniya/',
                'http://mohouse.ru/catalog/doma/arenda_predlozheniya/']
            dc = len(l)
            page = l[i]  
            oper = u'Аренда'
        else:
            break



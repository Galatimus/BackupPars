#!/usr/bin/python
# -*- coding: utf-8 -*-

from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
import os
import math
import json
from sub import conv
import xlsxwriter
import time
from datetime import datetime,timedelta
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab
import sys
reload(sys)
sys.setdefaultencoding('utf-8')



logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


    
i = 0
ls= open('links/mlsn_urls.txt').read().splitlines()
dc = len(ls)

places = []



while True:
    print '********************************************',i+1,'/',dc,'*******************************************'
    page = ls[i]
    lin = []
    class Mlsn_Urls(Spider):

        def prepare(self):
            self.f = page+'/prodazha-zemelnye-uchastki/'
            for p in range(1,50):
                try:
                    time.sleep(1)
                    g = Grab(timeout=10, connect_timeout=50)
                    g.proxylist.load_file(path='../tipa.txt',proxy_type='http') 
                    g.go(self.f)
                    print g.doc.code
                    self.num = re.sub('[^\d]', '',g.doc.select(u'//span[@class="long-label"]').text())
                    self.pag = int(math.ceil(float(int(self.num))/float(50)))
                    print self.num,self.pag
                    del g
                    break
                except(GrabTimeoutError,GrabNetworkError,IndexError,KeyError,AttributeError,GrabConnectionError):
                    print g.config['proxy'],'Change proxy'
                    g.change_proxy()
                    del g
                    continue            
            
        def task_generator(self):
            for x in range(1,self.pag+1):
                link = self.f+'?page=%s' % str(x)
                yield Task ('post',url=link,refresh_cache=True,network_try_count=100)

        def task_post(self,grab,task):
            for elem in grab.doc.select(u'//a[@class="location"]'):
                ur = grab.make_url_absolute(elem.attr('href')) 
                print ur
                lin.append(ur)

    bot = Mlsn_Urls(thread_number=5,network_try_limit=1000)
    bot.load_proxylist('../tipa.txt','text_file')
    bot.create_grab_instance(timeout=500, connect_timeout=500)    
    bot.run()
    print 'Save...' 
    print '***',len(lin),'****'
    time.sleep(2)    
    for item in lin:
        places.append(item)
    print 'Total...',len(places)
    time.sleep(1)
    try:
        i=i+1
        page = ls[i]
    except IndexError:
        break
print('*'*50)
print('Done Urls')
print('*'*50)
#liks = open('mlsn_zem.txt', 'w')
#for itm in places:
    #liks.write("%s\n" % itm)
#liks.close()
#print('Done')
time.sleep(5)
#os.system("/home/oleg/pars/mlsn/urlcom.py")
    
workbook = xlsxwriter.Workbook(u'zemm/0001-0002_00_У_001-0001_MLSN.xlsx')


oper = 'Продажа'


class Mlsn_Zem(Spider):
    def prepare(self):
        self.ws = workbook.add_worksheet()
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
        #l= open('mlsn_zem.txt').read().splitlines()
        self.dc = len(places)
        print self.dc
        for line in places:
            yield Task ('item',url=line,refresh_cache=True,network_try_count=100)




    def task_item(self, grab, task):
        try:
            sub = grab.doc.select(u'//title').text().split('MLSN.RU ')[1] 
        except IndexError:
            sub = ''	  
        try:
            ray = grab.doc.select(u'//h2[contains(text(),"Расположение")]/following-sibling::div').text()
        except IndexError:
            ray = ''          
        try:
            punkt= json.loads(grab.doc.rex_text(u'"locality":(.+?),"localityDistrict"'))['fullName']
        except (IndexError,TypeError):
            punkt = ''
        try:
            ter= grab.doc.rex_text(u'"fullName":"р-н.(.+?)"')
        except (IndexError,TypeError):
            ter =''
        try:
            ul = json.loads(grab.doc.rex_text(u'"street":(.+?),"house"'))
            uliza = ul['fullName']
        except (IndexError,TypeError):
            uliza = ''
        try:
            dom = re.sub('[^0-9\.]','',grab.doc.select(u'//div[contains(text(),"Обновлено")]').text())
        except IndexError:
            dom = ''



        try:
            metro = grab.doc.select(u'//td[contains(text(),"Категория использования")]/following-sibling::td').text()
            #print rayon
        except IndexError:
            metro = ''

        try:
            metro_min = grab.doc.select(u'//td[contains(text(),"Водопровод")]/following-sibling::td').text().replace(u'нет','')
            #print rayon
        except IndexError:
            metro_min = ''

        try:
            metro_tr = grab.doc.select(u'//div[@class="index__price"]/following-sibling::div').text()
        except IndexError:
            metro_tr = ''     

        try:
            price = grab.doc.select(u'//span[@class="Price__base"]').text()
            #print price + u' руб'	    
        except IndexError:
            price = ''     
        try:
            plosh_ob = grab.doc.select(u'//td[contains(text(),"Площадь участка")]/following-sibling::td').text()
            #print rayon
        except IndexError:
            plosh_ob = ''



        try:
            et = grab.doc.select(u'//td[contains(text(),"Газ")]/following-sibling::td').text().replace(u'нет','')
            #print price + u' руб'	    
        except IndexError:
            et = '' 

        try:
            etagn = grab.doc.select(u'//td[contains(text(),"Охрана")]/following-sibling::td').text().replace(u'нет','')
            #print price + u' руб'	    
        except IndexError:
            etagn = ''




        try:
            opis = grab.doc.select(u'//h2[contains(text(),"Описание")]/following-sibling::div/div').text() 
        except IndexError:
            opis = ''

        try:
            #id_ob = grab.doc.select(u'//span[@id="card_id"]').text()
            #url_phone = 'http://api.mlsn.ru/v1/'+vi+'/'+id_ob+'/viewContacts/'
            #r= requests.get(url_phone,timeout=1000)
            phone = grab.doc.rex_text(u'number":(.*?),')
        #print phone
        except IndexError:
            phone = ''

        try:
            lico = grab.doc.select(u'//div[@class="ContactsCard__contactName"]').text()
        except IndexError:
            lico = ''

        try:
            comp = grab.doc.select(u'//div[@class="ContactsCard__agencyName typography__bodySmall fonts__mainFont"][2]/span').text().replace(u'Агентство ','')
            #print rayon
        except IndexError:
            comp = ''

        try:
            data = re.sub('[^0-9\.]','',grab.doc.select(u'//div[contains(text(),"Добавлено")][2]').text())#.split(': ')[1])
        except IndexError:
            data = ''

        try:
            mesto = grab.doc.select(u'//h1/span[2]').text()
        except IndexError:
            mesto = ''
        try:
            price_sot = grab.doc.select(u'//div[@class="PropertyHeader__priceAddition typography__bodySmall fonts__mainFont"]').text()
        except IndexError:
            price_sot = ''		    




        subb = reduce(lambda sub, r: sub.replace(r[0], r[1]), conv, sub)





        projects = {'sub': subb,
                    'rayon': ray,
                    'punkt': punkt,
                    'teritor': ter,
                    'ulica': uliza,
                    'dom': dom,
                    'metro': metro,
                    'naz': metro_min,		           
                    'tran': metro_tr,
                    'cena': price,		           
                    'plosh_ob':plosh_ob,		           
                    'etach': et,
                    'etashost': etagn,      
                    'opis':opis,
                    'url':task.url,
                    'phone':re.sub(u'[^\d\-]','',phone),
                    'lico':lico,
                    'mesto':mesto,
                    'sot':price_sot,
                    'company':comp,
                    'data':data}



        yield Task('write',project=projects,grab=grab)






    def task_write(self,grab,task):

        print('*'*50)	       
        print  task.project['sub']
        print  task.project['rayon']
        print  task.project['punkt']
        print  task.project['teritor']
        print  task.project['ulica']
        print  task.project['dom']
        print  task.project['metro']
        print  task.project['naz']	      
        print  task.project['tran']
        print  task.project['cena']
        print  task.project['sot']
        print  task.project['plosh_ob']	       
        print  task.project['etach']
        print  task.project['etashost']
        print  task.project['opis']
        print  task.project['url']
        print  task.project['phone']
        print  task.project['lico']
        print  task.project['company']
        print  task.project['data']
        print  task.project['mesto']


        self.ws.write(self.result, 0,task.project['sub'])
        self.ws.write(self.result, 31,task.project['rayon'])
        self.ws.write(self.result, 2,task.project['punkt'])
        self.ws.write(self.result, 3,task.project['teritor'])
        self.ws.write(self.result, 4,task.project['ulica'])
        self.ws.write(self.result, 29,task.project['dom'])
        self.ws.write(self.result, 14,task.project['metro'])
        self.ws.write(self.result, 16,task.project['naz'])
        self.ws.write(self.result, 11,task.project['tran'])
        self.ws.write(self.result, 9,oper)
        self.ws.write(self.result, 10, task.project['cena'])
        #self.ws.write(self.result, 13, task.project['cena_m'])
        #self.ws.write(self.result, 14, task.project['col_komnat'])
        self.ws.write(self.result, 12, task.project['plosh_ob'])
        #self.ws.write(self.result, 16, task.project['plosh_gil'])
        self.ws.write(self.result, 11, task.project['sot'])
        self.ws.write(self.result, 31, task.project['mesto'])
        self.ws.write(self.result, 15, task.project['etach'])
        self.ws.write(self.result, 20, task.project['etashost'])
        self.ws.write(self.result, 22, task.project['opis'])
        self.ws.write(self.result, 23, u'MLSN.RU')
        self.ws.write_string(self.result, 24, task.project['url'])
        self.ws.write(self.result, 25, task.project['phone'])
        self.ws.write(self.result, 26, task.project['lico'])
        self.ws.write(self.result, 27, task.project['company'])
        self.ws.write(self.result, 28, task.project['data'])
        self.ws.write(self.result, 30, datetime.today().strftime('%d.%m.%Y'))


        print('*'*50)
        print 'Ready - '+str(self.result)+'/'+str(self.dc)
        logger.debug('Tasks - %s' % self.task_queue.size())
        print oper
        print('*'*50)
        self.result+= 1




        #if self.result > 50:
            #self.stop()



bot2 = Mlsn_Zem(thread_number=5,network_try_limit=1000)
bot2.load_proxylist('../tipa.txt','text_file')
bot2.create_grab_instance(timeout=50, connect_timeout=50)
bot2.run()
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
time.sleep(2)
workbook.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/small/mlsn_com.py")




#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
#from grab import Grab
import re
import os
import time
import xlsxwriter
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)




    

class Upn_com(Spider):
     def prepare(self):
	  self.workbook = xlsxwriter.Workbook(u'comm/0001-0062_00_C_001-0034_UPN-RU.xlsx')
	  self.ws = self.workbook.add_worksheet(u'Upn_Коммерческая')
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
	  for x in range(1,53):#92
               yield Task ('post',url='http://upn.ru/realty_offices_sale.htm?id=%d'%x,refresh_cache=True, network_try_count=100)
          for x1 in range(1,21):#92
	       yield Task ('post',url='http://upn.ru/realty_commercials_sale.htm?id=%d'%x1,refresh_cache=True, network_try_count=100)
	  for x2 in range(1,16):#92
	       yield Task ('post',url='http://upn.ru/realty_ready_business_sale.htm?id=%d'%x2,refresh_cache=True, network_try_count=100)
	  for x3 in range(1,8):#92
	       yield Task ('post',url='http://upn.ru/realty_industrials_sale.htm?id=%d'%x3,refresh_cache=True, network_try_count=100)
	  for x4 in range(1,5):#92
	       yield Task ('post',url='http://upn.ru/realty_stores_sale.htm?id=%d'%x4,refresh_cache=True, network_try_count=100)
	  for x5 in range(1,19):#92
	       yield Task ('post',url='http://upn.ru/realty_garages_sale.htm?id=%d'%x5,refresh_cache=True, network_try_count=100)
	  for x6 in range(1,49):#92
	       yield Task ('post',url='http://upn.ru/realty_offices_rent.htm?id=%d'%x6,refresh_cache=True, network_try_count=100)
	  for x7 in range(1,6):#92
	       yield Task ('post',url='http://upn.ru/realty_industrials_rent.htm?id=%d'%x7,refresh_cache=True, network_try_count=100)
	  for x8 in range(1,9):#92
	       yield Task ('post',url='http://upn.ru/realty_commercials_rent.htm?id=%d'%x8,refresh_cache=True, network_try_count=100)
	  for x9 in range(1,10):#92
	       yield Task ('post',url='http://upn.ru/realty_stores_rent.htm?id=%d'%x9,refresh_cache=True, network_try_count=100)
	       
	  yield Task ('post',url='http://upn.ru/realty_ready_business_rent.htm',refresh_cache=True, network_try_count=100)
	  yield Task ('post',url='http://upn.ru/realty_garages_rent.htm',refresh_cache=True, network_try_count=100) 
	       
          
	  
                 
        
        
            
            
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//tr[@class="robotno"]/td/a'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True, network_try_count=100,valid_status=(500,501,502))
	 
	 
     
     def task_item(self, grab, task):
	  if grab.doc.code == 200:
	       try:
		    mesto =  grab.doc.select(u'//b[contains(text(),"Адрес:")]/following::td[1]/a[1]').text()
	       except IndexError:
		    mesto =''
	       
	       try:
		    punkt = grab.doc.select(u'//b[contains(text(),"Адрес:")]/following::td[1]/a[1]').text().split(', ')[0]
	       except IndexError:
		    punkt = ''	       
	  
	       try:
		    ter =  grab.doc.select(u'//b[contains(text(),"Адрес:")]/following::td[1]/a[1]').text().split(', ')[1]
	       except IndexError:
		    ter =''
	       try:
		    uliza = grab.doc.select(u'//b[contains(text(),"Адрес:")]/following::td[1]/a[1]').text().split(', ')[2]
	       except IndexError:
		    uliza = ''
	       try:
		    dom = grab.doc.select(u'//div[@class="t13"]/a[4]').text()
	       except IndexError:
		    dom = ''
	  
	       try:
		    tip = grab.doc.select(u'//b[contains(text(),"Тип здания:")]/following::td[1]').text()
	       except IndexError:
		    tip = ''
	       try:
		    naz = grab.doc.select(u'//b[contains(text(),"Объект:")]/following::td[1]').text()
	       except IndexError:
		    naz =''
	       try:
		    klass =  grab.doc.select(u'//b[contains(text(),"Этажность:")]/following::td[1]').number()
	       except IndexError:
		    klass = ''
	       try:
		    price = grab.doc.select(u'//b[contains(text(),"Цена:")]/following::td[1]').text().replace(u' Отправить заявку на кредит','')
	       except IndexError:
		    price =''
	       try: 
		    try:
			 plosh = grab.doc.select(u'//b[contains(text(),"Общая площадь:")]/following::td[1]').text()
		    except IndexError:
			 plosh = grab.doc.select(u'//b[contains(text(),"Площадь:")]/following::td[1]').text()
	       except IndexError:
		    plosh=''
	       try:
		    ohrana = grab.doc.select(u'//b[contains(text(),"Этаж:")]/following::td[1]').number()
	       except IndexError:
		    ohrana =''
	       try:
		    gaz =  re.findall(">(.*?)<",re.sub('\s+',' ',grab.response.body.split('Материал стен:')[1].split('\n')[1]))[0]
	       except IndexError:
		    gaz =''
	       try:
		    voda = grab.doc.select(u'//b[contains(text(),"Год постройки:")]/following::td[1]').text()
	       except IndexError:
		    voda =''
	       try:
		    kanal = re.findall(">(.*?)<",re.sub('\s+',' ',grab.response.body.split('Высота потолков:')[1].split('\n')[1]))[0]
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
		    opis = grab.doc.select(u'//tr[contains(@id,"divDescription")]').text()
	       except IndexError:
		    opis = ''
	       try:
		    try:
			 lico = grab.doc.select(u'//div[@class="agent-contact-info"]/div/h3').text()
		    except IndexError: 
			 lico = grab.doc.select(u'//dt[contains(text(),"Агент")]/following-sibling::dd[1]').text()
	       except IndexError:
		    lico = ''
	       try:
		    comp = grab.doc.select(u'//b[contains(text(),"Агентство недвижимости:")]/following::td[1]/a').attr('title')
	       except IndexError:
		    comp = ''
	       try:
		    data1 = grab.doc.select(u'//dt[contains(text(),"Обновленно")]/following-sibling::dd[1]').text() 
	       except IndexError:   
		    data1 = ''
	       try: 
		    data = grab.doc.select(u'//h1[2]').text()
	       except IndexError:
		    data=''
	  
	       try:
		    try:
			 phone = grab.doc.select(u'//b[contains(text(),"Телефон агентства:")]/following::td[1]').text()
		    except IndexError:
			 phone = grab.doc.select(u'//b[contains(text(),"Телефон агента:")]/following::td[1]').text()
	       except IndexError:
		    phone = ''
		    
	       try:
		    if 'sale' in task.url:
			 oper = u'Продажа'
		    else:
			 oper = u'Аренда'
	       except IndexError:
		    oper = ''
	       
	       
	       
	       
	       
	       projects = {'sub': 'Свердловская область',
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
		           'opis': opis,
		           'url': task.url,
		           'phone': re.sub('[^\d\,]','',phone),
		           'lico':lico,
		           'company': comp,
		           'data':data,
		           'data1':data1,
		           'oper':oper}
	       
	       
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
	  print  task.project['data1']
	  
	  
	  
	  
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 24, task.project['adress'])
	  self.ws.write(self.result, 3, task.project['terit'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 7, task.project['dom'])
	  self.ws.write(self.result, 8, task.project['tip'])
	  self.ws.write(self.result, 9, task.project['naz'])
	  self.ws.write(self.result, 16, task.project['klass'])
	  self.ws.write(self.result, 11, task.project['cena'])
	  self.ws.write(self.result, 14, task.project['plosh'])
	  self.ws.write(self.result, 15, task.project['ohrana'])
	  #self.ws.write(self.result, 16, task.project['gaz'])
	  #self.ws.write(self.result, 17, task.project['voda'])
	  #self.ws.write(self.result, 17, task.project['kanaliz'])
	  #self.ws.write(self.result, 23, task.project['electr'])
	  #self.ws.write(self.result, 24, task.project['teplo'])
	  self.ws.write(self.result, 18, task.project['opis'])
	  self.ws.write(self.result, 19, u'Уральская палата недвижимости')
	  self.ws.write_string(self.result, 20, task.project['url'])
	  self.ws.write(self.result, 21, task.project['phone'])
	  #self.ws.write(self.result, 29, task.project['lico'])
	  self.ws.write(self.result, 23, task.project['company'])
	  self.ws.write(self.result, 33, task.project['data'])
	  #self.ws.write(self.result, 32, task.project['data1'])
	  self.ws.write(self.result, 31, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write(self.result, 28, task.project['oper'])
	  print('*'*100)
	  #print self.sub
	  print 'Ready - '+str(self.result)
	  logger.debug('Tasks - %s' % self.task_queue.size()) 
	  print  task.project['oper']

	  print('*'*100)
	  self.result+= 1





	  #if self.result > 20:
	       #self.stop() 
	  
	 

     
bot = Upn_com(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5000, connect_timeout=50000)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')
try:
     command = 'mount -a'
     os.system('echo %s|sudo -S %s' % ('1122', command))
     time.sleep(5)
     bot.workbook.close()
     print('Done')
except IOError:
     time.sleep(30)
     os.system('echo %s|sudo -S %s' % ('1122', 'mount -a'))
     time.sleep(10)
     bot.workbook.close()
     print('Done!')
print('Done!')

time.sleep(5)
os.system("/home/oleg/pars/small/zdanie.py")






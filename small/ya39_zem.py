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


workbook = xlsxwriter.Workbook(u'zemm/0001-0080_00_У_001-0073_YA39.xlsx')

    

class Cian_Zem(Spider):
     def prepare(self):
	  #self.f = page
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
	  self.ws.write(0, 29, u"ДАТА_ПАРСИНГА")
	  self.ws.write(0, 30, u"МЕСТОПОЛОЖЕНИЕ")
	       
	       
	  self.result= 1
	
	       
    
     def task_generator(self):
	  for x in range(1,36):#30
               yield Task ('post',url='http://ya39.ru/realty/parcel/?action_id=3&PAGEN_1=%d'%x,network_try_count=100)
          yield Task ('post',url='http://ya39.ru/realty/parcel/?action_id=4',network_try_count=100)
                 
        
        
            
            
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//a[@class="preTit"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur
	       yield Task('item', url=ur,refresh_cache=True,network_try_count=100)
	 
        
        
     def task_item(self, grab, task):
	  try:
	       sub = u'Калиниградская область'
	  except DataNotFound:
	       sub = ''
	  try:
	       ray = grab.doc.select(u'//em/a[contains(text(),"р-н")]').text()
	     #print ray 
	  except DataNotFound:
	       ray = ''          
	  try:
	       punkt =grab.doc.select(u'//div[contains(text(),"Адрес:")]/following-sibling::div').text().split(u', ')[0].replace(sub,'')
	  except IndexError:
	       punkt = ''
	       
	  try:
	       #if  grab.doc.select(u'//h1[@class="object_descr_addr"]/a[contains(text()," район")]').exists()==True:
	       ter= grab.doc.select(u'//div[contains(text(),"Адрес:")]/following-sibling::div').text().split(u', ')[1].replace(sub,'').replace(punkt,'')
	       #else:
		    #ter= ''#grab.doc.select(u'//h1[@class="object_descr_addr"]').text().split(', ')[2]
	  except IndexError:
	       ter =''
	       
	  try:
	       li = grab.doc.select(u'//div[contains(text(),"Адрес:")]/following-sibling::div').text().split(u', ')[2].replace(sub,'').replace(punkt,'').replace(ter,'').replace(',','')
	       if 'г. ' in li:
	       	    punkt = li
		    uliza = ''
	       elif ' п.' in li:
		    punkt = li
		    uliza = '' 
	       else:
		    uliza = li
	        
	  except IndexError:
	       uliza = ''
	  try:
	       dm = grab.doc.select(u'//div[contains(text(),"Адрес:")]/following-sibling::div').text()
               dom = re.compile(r'[0-9]+$',re.M).search(dm).group(0)	 
	  except (IndexError,AttributeError):
	       dom = ''
	       
	  try:
	       trassa = grab.doc.select(u'//div[contains(text(),"Адрес:")]/following-sibling::div').text()
		#print rayon
	  except DataNotFound:
	       trassa = ''
	       
	  try:
	       udal = grab.doc.select(u'//td[contains(text(),"Класс:")]/following-sibling::td').text()
	  except DataNotFound:
	       udal = ''
	       
	  try:
	       price = re.sub('[^\d]','',grab.doc.select(u'//h1/span[2]').text())+u' р.'
	  except DataNotFound:
	       price = ''
	       
	  
	       
	  try:
	       plosh = grab.doc.select(u'//div[contains(text(),"Площадь участка:")]/following-sibling::div').text()
	  except DataNotFound:
	       plosh = ''

	  try:
	       vid = grab.doc.select(u'//div[contains(text(),"Категория земли:")]/following-sibling::div[1]').text()
	  except DataNotFound:
	       vid = '' 
	       
	       
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
	  except IndexError:
	       teplo =''
	       
	  try:
	       oper = u'Продажа' 
	  except IndexError:
	       oper = ''               
	      
		    
	  try:
	       opis = grab.doc.select(u'//h3[contains(text(),"Описание:")]/following-sibling::p').text() 
	  except IndexError:
	       opis = ''
	       
	  try:
	       phone = re.sub('[^\d]', '',grab.doc.select(u'//div[@class="phon"]').text())
	  except IndexError:
	       phone = ''
	       
	  try:
	       lc = grab.doc.select(u'//div[@class="col-xs-9 col-sm-9 col-md-9 col-lg-9"]/div').text()
	       if '-' in lc:
		    comp=lc
		    lico = ''
	       else:
		    lico=lc
		    comp = ''
	  except IndexError:
	       lico = ''
	       comp = ''
	       
	  try:
	       con = [ (u' августа ',u'.08.'), (u' июля ',u'.07.'),
	            (u' мая ',u'.05.'),(u' июня ',u'.06.'),
	            (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
	            (u' января ',u'.01.'),(u' декабря ',u'.12.'),
	            (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
	            (u' февраля ',u'.02.'),(u' октября ',u'.10.')]
	       d = grab.doc.select(u'//span[@class="glyphicon glyphicon-time"]/following-sibling::text()').text()#.split(u' добавлено ')[1]
	       if 'мин. назад' in d:		    
		    data = datetime.today().strftime('%d.%m.%Y')
	       elif 'Вчера' in d:
		    data = '{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1))
	       elif 'Сегодня' in d:
		    data = (datetime.today().strftime('%d.%m.%Y'))
	       else:
		    if '2018' in d:
	                 data = reduce(lambda d, r: d.replace(r[0], r[1]), con, d).split(':')[0][:-2]+'2018'
		    else:
			 data = reduce(lambda d, r: d.replace(r[0], r[1]), con, d).split(':')[0][:-2]+'2019'
	  except IndexError:
	       data = ''
		    
	  
						   
	       
	  projects = {'url': task.url,
                      'sub': sub,
                      'rayon': ray,
                      'punkt': punkt,
                      'teritor': ter,
                      'ulica': uliza,
	              'dom': dom,
                      'trassa': trassa,
                      'udal': udal,
                      'cena': price,
                      'plosh':plosh,
                      'vid': vid,
                      'ohrana':ohrana,
                      'gaz': gaz,
                      'voda': voda,
                      'kanaliz': kanal,
                      'electr': elek,
                      'teplo': teplo,
                      'opis':opis,
                      'phone':phone,
                      'lico':lico,
                      'company':comp,
                      'data':data[:10],
                      'oper':oper
                      }
          
	  yield Task('write',project=projects,grab=grab,refresh_cache=True)
            
     def task_write(self,grab,task):
	  print('*'*50)
	  print  task.project['sub']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['teritor']
	  print  task.project['ulica']
	  print  task.project['dom']
	  print  task.project['trassa']
	  print  task.project['udal']
	  print  task.project['cena']
	  print  task.project['plosh']
	  print  task.project['vid']
	  print  task.project['ohrana']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanaliz']
	  print  task.project['electr']
	  print  task.project['teplo']
	  print  task.project['opis']
	  print task.project['url']
	  print  task.project['phone']
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['data']
	  print  task.project['oper']
	  
	  #global result
	  self.ws.write(self.result, 0, task.project['sub'])
	  self.ws.write(self.result, 1, task.project['rayon'])
	  self.ws.write(self.result, 2, task.project['punkt'])
	  self.ws.write(self.result, 3, task.project['teritor'])
	  self.ws.write(self.result, 4, task.project['ulica'])
	  self.ws.write(self.result, 5, task.project['dom'])
	  self.ws.write(self.result, 30, task.project['trassa'])
	  self.ws.write(self.result, 14, task.project['udal'])
	  self.ws.write(self.result, 9, task.project['oper'])
	  self.ws.write(self.result, 10, task.project['cena'])
	  self.ws.write(self.result, 12, task.project['plosh'])
	  self.ws.write(self.result, 13, task.project['vid'])
	  self.ws.write(self.result, 15, task.project['gaz'])
	  self.ws.write(self.result, 16, task.project['voda'])
	  self.ws.write(self.result, 17, task.project['kanaliz'])
	  self.ws.write(self.result, 18, task.project['electr'])
	  self.ws.write(self.result, 19, task.project['teplo'])
	  self.ws.write(self.result, 20, task.project['ohrana'])	       
	  self.ws.write(self.result, 22, task.project['opis'])
	  self.ws.write(self.result, 23, u'Издательский дом "Ярмарка"')
	  self.ws.write_string(self.result, 24, task.project['url'])
	  self.ws.write(self.result, 25, task.project['phone'])
	  self.ws.write(self.result, 26, task.project['lico'])
	  self.ws.write(self.result, 27, task.project['company'])
	  self.ws.write(self.result, 28, task.project['data'])
	  self.ws.write(self.result, 29, datetime.today().strftime('%d.%m.%Y'))
	  print('*'*50)
	  #print task.sub
	  
	  print 'Ready - '+str(self.result)#+'/'+task.project['koll']
	  logger.debug('Tasks - %s' % self.task_queue.size())
	  #print '*',i+1,'/',dc,'*'
	  #print oper
	  print('*'*50)	       
	  self.result+= 1
	       
	 
	  #if self.result > 50:
	       #self.stop()	            
     
bot = Cian_Zem(thread_number=5,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5, connect_timeout=5)
bot.run()
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
sudoPassword = '1122'
command = 'mount -a'# cifs //192.168.1.6/d /home/oleg/pars -o username=oleg,password=1122,_netdev,rw,iocharset=utf8,file_mode=0777,dir_mode=0777'
p = os.system('echo %s|sudo -S %s' % (sudoPassword, command))
print p
time.sleep(2)
workbook.close()
print('Done!')

time.sleep(5)
os.system("/home/oleg/pars/small/ya39_com.py")






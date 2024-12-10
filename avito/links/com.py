#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import math
import random
from datetime import datetime,timedelta
import xlsxwriter
from cStringIO import StringIO
import pytesseract
from PIL import Image 
import os
import time
import base64
from grab import Grab
import sys
reload(sys)
sys.setdefaultencoding('utf-8')



logging.basicConfig(level=logging.DEBUG)



class Avito(Spider):


     def prepare(self):
	  self.workbook = xlsxwriter.Workbook(u'Avito_Коммерческая_цены.xlsx')
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
	  self.g = 0 
	  self.sub =''

     def task_generator(self):
	  for line in open('avito.txt').read().splitlines():
               yield Task ('item',url=line.strip(),refresh_cache=True,network_try_count=100)
	  
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//a[@class="item-description-title-link"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur	      
	       yield Task('item',url=ur,refresh_cache=True,network_try_count=100)
	     
     def task_item(self, grab, task):
	  
	  try:
	       lin = []
	       for em in grab.doc.select(u'//div[@class="item-map-location"]/span[@itemprop="name"]'):
		    urr = em.text()
		    lin.append(urr)
	       mesto = ",".join(lin).replace(u'Адрес:,','')+','+grab.doc.select(u'//span[@class="item-map-address"]').text()
	  except IndexError:
	       mesto = ''
     
	  try:
	       ray =  grab.doc.select(u'//span[@class="item-map-address"]/span[contains(text(), "р-н")]/text()').text().replace(',','')
	  except IndexError:
	       ray = ''
	  try:
	       if self.sub == u'Москва':
		    punkt= u'Москва'
	       elif self.sub == u'Санкт-Петербург':
		    punkt= u'Санкт-Петербург'
	       elif self.sub == u'Севастополь':
		    punkt= u'Севастополь'
	       else:		    
		    punkt = grab.doc.rex_text(u'selected >(.*?)</option>')
	  except IndexError:
	       punkt = ''
	  try:
	       uliza = grab.doc.select(u'//span[@itemprop="streetAddress"]').text()
	       #uliza = re.split('(\W+)',ul)[1]
	  except IndexError:
	       uliza = ''
	  try:
	       dom = u'Коммерческая недвижимость'#grab.doc.select(u'//span[@itemprop="streetAddress"]').text()
	       #dom =re.split('\W+', d,1)[1]
	  except IndexError:
	       dom = ''
	  
	  try:
	       tip = grab.doc.select(u'//div[@class="item-price-old"]').text()
	  except IndexError:
	       tip = ''
	  try:
	       naz = grab.doc.select(u'//div[@class="b-catalog-breadcrumbs"]/a[5]').text()#.split(',')[0].replace(tip,'').replace(u'Продается ','').replace(u'Продам ','')
	  except IndexError:
	       naz = ''
	  try:
	       klass = grab.doc.select(u'//span[contains(text(),"Класс здания:")]/following-sibling::text()').text()
	  except IndexError:
	       klass = ''
	  try:
	       price = grab.doc.select('//span[@class="price-value-string js-price-value-string"]').text()
	  except IndexError:
	       price = ''
	  try:
	       et = ''#grab.doc.select(u'//ul[@id="flat_data"]/li[contains(text(),"этаж")]').number()
	  except IndexError:
	       et = ''
     
	  try:
	       et2 = grab.doc.select(u'//li[contains(text(),"этажность")]').number()
	  except IndexError:
	       et2 = ''
	       
	  try:
	       god = grab.doc.select(u'//li[contains(text(),"год постройки")]/b').number()
	  except IndexError:
	       god = ''
	  
	  try:
	       mat = grab.doc.select(u'//div[@class="item-price-sub-price"]').text()#.replace(u'за м² ','')
	  except IndexError:
	       mat = ''

	  try:
	       pot = grab.doc.select(u'//div[@class="seller-info-label"][contains(text(),"Адрес")]/following-sibling::div').text()
	  except IndexError:
	       pot = ''

	  try:
	       sos = grab.doc.select(u'//span[@class="item-map-metro"]').text().split(u' (')[0]
	  except IndexError:
	       sos = ''
	       
	  try:
	       plosh = grab.doc.select(u'//span[contains(text(),"Площадь:")]/following-sibling::text()').text()
	  except IndexError:
	       plosh = ''
	  
	  try:
	       gaz = grab.doc.select(u'//span[@class="item-map-metro"]').text().split(' (')[1].replace(')','')
	  except IndexError:
	       gaz =''
	  try:
	       voda = grab.doc.select(u'//h1').text()
	  except IndexError:
	       voda =''
	  try:
	       kanal = grab.doc.rex_text(u'data-map-lat="(.*?)"')
	  except IndexError:
	       kanal =''
	  try:
	       elekt = grab.doc.rex_text(u'data-map-lon="(.*?)"')
	  except IndexError:
	       elekt =''
	  try:
	       teplo = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"отопление")]').text().replace(u'есть отопление',u'есть').replace(u'нет отопления','')
	  except IndexError:
	       teplo =''
	  try:
	       ohrana = grab.doc.select(u'//ul[@id="infra_data"]/li[contains(text(),"охрана")]').text().replace(u'есть охрана',u'есть').replace(u'нет охраны','')
	  except IndexError:
	       ohrana =''
	  try:
	       opis = grab.doc.select(u'//div[@class="item-description"]/div').text() 
	  except IndexError:
	       opis = ''
	  try:
	       try:
		    try:
			 lico = grab.doc.select(u'//div[contains(text(),"Продавец")]/following-sibling::div/div[1]').text()
		    except IndexError:
			 lico = grab.doc.select(u'//div[contains(text(),"Арендодатель")]/following-sibling::div/div[1]').text()
	       except IndexError:
		    lico = grab.doc.select(u'//div[contains(text(),"Контактное лицо")]/following-sibling::div').text()
	  except IndexError:
	       lico = ''
	  
	  try:
	       com = grab.doc.select(u'//div[contains(text(),"Агентство")]/following-sibling::div/div[1]').text()
	  except IndexError:
	       com = ''
	       
	  try:
	       rphone = re.sub('[^\d]','',grab.doc.select(u'//span[@class="item-phone-button-sub-text"]').text()+str(random.randint(1000000,9999999)))
	  except IndexError:
	       rphone = ''			    
	  try:
	       conv = [ (u'сегодня', (datetime.today().strftime('%d.%m.%Y'))),
	            (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1))),
	            (u'июня', '.06.2017'),(u'июля', '.07.2016'),(u'августа', '.08.2016'),(u'января', '.01.2017'),(u'февраля', '.02.2017'),
	            (u'марта', '.03.2017'),(u'апреля', '.04.2017'),(u'мая', '.05.2017'),
	            (u'ноября', '.11.2016'),(u'сентября', '.09.2016'),(u'октября', '.10.2016'),(u'декабря', '.12.2016')]
	       dt= grab.doc.rex_text(u'размещено (.*?) в')
	       data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt).replace(' ','')
	  except IndexError:
	       data = ''		    
	  
	  
	  
	  
	  
	  
	  projects = {'url': task.url,
                    'sub': self.sub,
                    'rayon': ray,
                    'punkt': punkt,
                    'ulica': uliza,
                    'dom': dom,
                    'naz': naz,
                    'tip': tip,
                    'price': price,
                    'klass': klass,
                    'ploshad': plosh,
                    'et': et,
                    'ets': et2,
                    'god': god,
                    'mat': mat,
                    'potolok': pot,
                    'sost': sos,
                    'gaz': gaz,
                    'voda':voda,
                    'kanal': kanal,
                    'elekt': elekt,
                    'teplo': teplo,
                    'ohrana': ohrana,
                    'opis': opis,
                    'mesto':mesto,
                    'lico':lico,
                    'company':com,
                    'phone2':rphone,
                    'data': data }
	  
	  #try:
	       ##ad_id= re.sub(u'[^\d]','',task.url[-9:])
	       #ad_id= re.sub(u'[^\d]','',grab.doc.rex_text(u'data-item-id="(.*?)"'))
	       #ad_phone = re.sub(u'[^0-9a-z]','',grab.doc.rex_text(u'avito.item.phone(.*?);'))
	       #ad_subhash = re.findall(r'[0-9a-f]+', ad_phone)
	       #if int(ad_id) % 2 == 0:
		    #ad_subhash.reverse()
	       #ad_subhash = ''.join(ad_subhash)[::3]
	       #link = 'https://www.avito.ru/items/phone/'+ad_id+'?pkey='+ad_subhash
	       #headers ={'Accept': 'application/json, text/javascript, */*; q=0.01',
	                      #'Accept-Encoding': 'gzip,deflate',
	                      #'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
	                      #'Cookie': 'sessid='+ad_id+'.'+ad_subhash,
	                      #'Host': 'www.avito.ru',
	                      #'Referer': task.url,
	                      #'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0', 
	                      #'X-Requested-With' : 'XMLHttpRequest'}
	       #gr = Grab()
	       #gr.setup(url=link,headers=headers)
	       #yield Task('phone',grab=gr,project=projects,refresh_cache=True,network_try_count=100)
	  #except IndexError:
	       #yield Task('phone',grab=grab,project=projects)
	       
     #def task_phone(self, grab, task):
	  #try:
	       #data_image64 = grab.response.json['image64'].replace('data:image/png;base64,','') 
	       #imgdata = base64.b64decode(data_image64)
	       #im = Image.open(StringIO(imgdata))
	       #x,y = im.size
	       #phon = pytesseract.image_to_string(im.convert("RGB").resize((int(x*2), int(y*3)),Image.BICUBIC))
	  #except (IndexError,ValueError,GrabNetworkError,GrabTimeoutError,IOError):
	       #phon = ''
	       
	  #phone=re.sub(u'[^\d]','',phon)
	  #if phone == '05':
	       #phone = task.project['phone2']		    
	  
	  #projects = {'phone': re.sub(u'[^\d]','',phone)}
		      

		 
	  
	  yield Task('write',project=projects,grab=grab)
	  
	  
	  
	  
	  
     def task_write(self,grab,task):
	  #if task.phone <> '':    
	  print('*'*100)
	  print  task.project['sub']
	  print  task.project['rayon']
	  print  task.project['punkt']
	  print  task.project['ulica']
	  print  task.project['dom']
	  print  task.project['naz']
	  print  task.project['tip']
	  print  task.project['price']
	  print  task.project['klass']
	  print  task.project['ploshad']
	  print  task.project['et']
	  print  task.project['ets']
	  print  task.project['god']
	  print  task.project['mat']
	  print  task.project['potolok']
	  print  task.project['sost']
	  print  task.project['gaz']
	  print  task.project['voda']
	  print  task.project['kanal']
	  print  task.project['elekt']
	  #print  task.project['teplo']
	  #print  task.project['ohrana']
	  print  task.project['opis']
	  print  task.project['url']
	  #print  task.phone
	  print  task.project['lico']
	  print  task.project['company']
	  print  task.project['mesto']
	  print  task.project['data']
	  
     
	  self.ws.write(self.result,0, task.project['sub'])
	  self.ws.write(self.result,3, task.project['rayon'])
	  self.ws.write(self.result,2, task.project['punkt'])
	  self.ws.write(self.result,4, task.project['ulica'])
	  self.ws.write(self.result,7, task.project['dom'])
	  self.ws.write(self.result,9, task.project['naz'])
	  self.ws.write(self.result,12, task.project['tip'])
	  #self.ws.write(self.result,28, oper)
	  self.ws.write(self.result,11, task.project['price'])
	  self.ws.write(self.result,10, task.project['klass'])
	  self.ws.write(self.result,14, task.project['ploshad'])
	  self.ws.write(self.result,15, task.project['et'])
	  self.ws.write(self.result,16, task.project['ets'])
	  self.ws.write(self.result,17, task.project['god'])
	  self.ws.write(self.result,13, task.project['mat'])
	  self.ws.write(self.result,25, task.project['potolok'])
	  self.ws.write(self.result,26, task.project['sost'])
	  self.ws.write(self.result,27, task.project['gaz'])
	  self.ws.write(self.result,33, task.project['voda'])
	  self.ws.write(self.result,34, task.project['kanal'])
	  self.ws.write(self.result,35, task.project['elekt'])
	  #self.ws.write(self.result,24, task.project['teplo'])
	  #self.ws.write(self.result,19, task.project['ohrana'])
	  self.ws.write(self.result,18, task.project['opis'])
	  self.ws.write(self.result,19, u'AVITO.RU')
	  self.ws.write_string(self.result,20, task.project['url'])
	  #self.ws.write(self.result,21, task.phone)
	  self.ws.write(self.result,22, task.project['lico'])
	  self.ws.write(self.result,23, task.project['company'])
	  self.ws.write(self.result,29, task.project['data'])
	  self.ws.write(self.result,31, datetime.today().strftime('%d.%m.%Y'))
	  #self.ws.write(self.result,34, task.project['data1'])
	  self.ws.write(self.result,24, task.project['mesto'])
	 
	  
	  print('*'*100)
	  print self.sub	       
	  print 'Ready - '+str(self.result)
	  print 'Tasks - %s' % self.task_queue.size()
	  #print oper
	  print('*'*100)
	 
	  self.result+= 1
	  
	  #if self.result >10:
	       #self.stop()
		    
	       #if self.result >= self.num:
		    #self.stop()		  

bot = Avito(thread_number=5, network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file',proxy_type='http')
bot.create_grab_instance(timeout=5000, connect_timeout=5000)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')    
command = 'mount -t cifs //192.168.1.6/d /home/oleg/pars -o username=oleg,password=1122,iocharset=utf8,file_mode=0777,dir_mode=0777' #'mount -a -O _netdev'
#command = 'apt autoremove'
p = os.system('echo %s|sudo -S %s' % ('1122', command))
print p
time.sleep(1)
bot.workbook.close()
#workbook.close()
print('Done!')

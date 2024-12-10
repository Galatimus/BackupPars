#!/usr/bin/python
# -*- coding: utf-8 -*-

import logging
import time
import re
#import random
import dryscrape
from datetime import datetime,timedelta
#from sub import conv
import webkit_server
import random
#import socket
from xlsxwriter import Workbook
from dryscrape.mixins import WaitTimeoutError 
from webkit_server import InvalidResponseError,EndOfStreamError,NoX11Error
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
logging.basicConfig(level=logging.DEBUG)



#dryscrape.start_xvfb()
#sess = dryscrape.Session()
#sess.set_header('user-agent', 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:52.0) Gecko/20100101 Firefox/52.0')
#sess.set_attribute('auto_load_images', False)
##sess.set_error_tolerant(True) 


workbook = Workbook('0001-0002_00_C_001-0217_BCINF.xlsx')


ws = workbook.add_worksheet()
ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
ws.write(0, 4, u"УЛИЦА")
ws.write(0, 5, u"ДОМ")
ws.write(0, 6, u"ОРИЕНТИР")
ws.write(0, 7, u"СЕГМЕНТ")
ws.write(0, 8, u"ТИП_ПОСТРОЙКИ")
ws.write(0, 9, u"НАЗНАЧЕНИЕ_ОБЪЕКТА")
ws.write(0, 10, u"ПОТРЕБИТЕЛЬСКИЙ_КЛАСС")
ws.write(0, 11, u"СТОИМОСТЬ")
ws.write(0, 12, u"ИЗМЕНЕНИЕ_СТОИМОСТИ")
ws.write(0, 13, u"ДОПОЛНИТЕЛЬНЫЕ_КОММЕРЧЕСКИЕ_УСЛОВИЯ")
ws.write(0, 14, u"ПЛОЩАДЬ")
ws.write(0, 15, u"ЭТАЖ")
ws.write(0, 16, u"ЭТАЖНОСТЬ")
ws.write(0, 17, u"ГОД_ПОСТРОЙКИ")
ws.write(0, 18, u"ОПИСАНИЕ")
ws.write(0, 19, u"ИСТОЧНИК_ИНФОРМАЦИИ")
ws.write(0, 20, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
ws.write(0, 21, u"ТЕЛЕФОН")
ws.write(0, 22, u"КОНТАКТНОЕ_ЛИЦО")
ws.write(0, 23, u"КОМПАНИЯ")
ws.write(0, 24, u"МЕСТОПОЛОЖЕНИЕ")
ws.write(0, 25, u"МЕСТОРАСПОЛОЖЕНИЕ")
ws.write(0, 26, u"БЛИЖАЙШАЯ_СТАНЦИЯ_МЕТРО")
ws.write(0, 27, u"РАССТОЯНИЕ_ДО_БЛИЖАЙШЕЙ_СТАНЦИИ_МЕТРО")
ws.write(0, 28, u"ОПЕРАЦИЯ")
ws.write(0, 29, u"ДАТА_РАЗМЕЩЕНИЯ")
ws.write(0, 30, u"ДАТА_ОБНОВЛЕНИЯ")
ws.write(0, 31, u"ДАТА_ПАРСИНГА")
ws.write(0, 32, u"КАДАСТРОВЫЙ_НОМЕР")
ws.write(0, 33, u"ЗАГОЛОВОК")
ws.write(0, 34, u"ШИРОТА_ИСХ")
ws.write(0, 35, u"ДОЛГОТА_ИСХ")
result= 1           
           


z = 0
lin= open('Links/biz.txt').read().splitlines()

dryscrape.start_xvfb()
server = webkit_server.Server()
server_conn = webkit_server.ServerConnection(server=server)
driver = dryscrape.driver.webkit.Driver(connection=server_conn)
sess = dryscrape.Session(driver=driver)


try:           
           while True: 
                      print z+1,'/',str(len(lin))
                      proxy = random.choice(list(open('../tipa.txt').read().splitlines())).split(':')[0]
                      print proxy                     
                      try:
                                 print lin[z]
                      except IndexError:
                                 workbook.close()
                                 print('Done')                                 
                                 break
                      try:
                                 sess.set_timeout(30)
                                 
                                 #sess.set_header('Host', 'www.cian.ru')
                                 sess.set_header('User-Agent', 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0')
                                 sess.set_header('Accept', 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8')
                                 sess.set_header('Accept-Language', 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3')
                                 #sess.set_header('Accept-Encoding', 'gzip, deflate, br')
                                 sess.set_proxy(host = proxy, port = 4045)
                                 sess.set_cookie('_CIAN_GK=728644fe-916e-416e-af6d-41d9bf52f7cf')
                                 #sess.set_proxy(host= proxy, port=8080, user='Ivan', password='tempuvefy')
                                 sess.visit(lin[z])
                                 #sess.wait_for(lambda: sess.at_xpath(u'//span[contains(text(),"Показать телефон")]'))
                                 print sess.status_code()
                      except BaseException as e:
                                 print str(e)
                                 if 'wait_for timed out'in e:
                                            time.sleep(2)
                                            z=z+1
                                            continue
                      if 'captcha'in sess.url():    
                                 time.sleep(2)
                                 continue
                      #except WaitTimeoutError:
                                 #print 'WaitTimeoutError'
                                 ##sess.reset()
                                 #z=z+1
                                 #continue                                 
                      try:
                                 ray1 = sess.at_xpath(u'//title').text()
                                 ray = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", ray1)
                                 ray = re.sub(u"[.,\-\s]{3,}", " ", ray)                                 
                      except :
                                 ray = ''
                                 
                      print sess.url()
                                 
                      for link in sess.xpath('//a[contains(@href,"sale/commercial")]'):
                                 url = link['href']
                                 print url                                 
                      #try:
                                 #if 'moscow' in lin[z]:
                                            #punkt = u'Москва'
                                 #else:
                                            #punkt=sess.at_xpath(u'//ul[@class="breadcrumbs"]/li[2]/span/a/span[1]').text()
                      #except :
                                 #punkt = ''                                            
                      #try:
                                 #sub = reduce(lambda punkt, r: punkt.replace(r[0], r[1]), conv, punkt)
                      #except :
                                 #sub =''
                      #try:
                                 #uliza = sess.at_xpath(u'//div[@class="street"]/a').text()
                      #except :
                                 #uliza = ''
                                 
                      #try:
                                 #dom1 = sess.at_xpath(u'//div[@class="metro"]/a/p').text().split(' (')[0]
                                 #dom = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", dom1)
                                 #dom = re.sub(u"[.,\-\s]{3,}", " ", dom)                                 
                      #except :
                                 #dom = ''                                                       
                                 
                      #try:
                                 #seg1 = sess.at_xpath(u'//span[@class="left"]/a').text()
                                 #seg = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", seg1)
                                 #seg = re.sub(u"[.,\-\s]{3,}", " ", seg)                                 
                      #except :
                                 #seg = ''                                            
                      #try:
                                 #naz = sess.at_xpath(u'//tr[@class="object-id"]/td[2]').text()
                      #except :
                                 #naz = ''                                                       
                      #try:
                                 #price = sess.at_xpath(u'//meta[@name="keywords"]')['content'].split(', ')[1]
                      #except :
                                 #price = ''
                                 
                      #try:
                                 #cena_za = sess.at_xpath(u'//meta[@name="description"]')['content'].split(': ')[1].split(' ')[0]
                      #except :
                                 #cena_za = ''
                      #try: 
                                 #klass = sess.at_xpath(u'//span[@class="white-text badge"]').text()
                      #except :
                                 #klass =''
                      #try:
                                 #plosh = sess.at_xpath(u'//meta[@name="keywords"]')['content'].split(', ')[0]
                      #except :
                                 #plosh = ''
                                 
                      #try:
                                 #et = sess.at_xpath(u'//tr[@class="object-id"]/td[3]').text()
                      #except :
                                 #et = ''
                      
                      #try:
                                 #et2 = sess.at_xpath(u'//title').text()
                      #except :
                                 #et2 = ''
                                 
                      #try:
                                 #god = sess.at_xpath(u'//div[@class="metro"]/a/p').text().split(' (')[1].replace(')','')
                      #except :
                                 #god =''
                                 
                      #try:
                                 #zag = sess.at_xpath(u'//span[contains(text(),"Площадь комнат")]/following-sibling::span[1]').text()
                      #except :
                                 #zag =''
                                 
                      #try:
                                 #do_m = sess.at_xpath(u'//span[contains(text(),"Вид из окон")]/following-sibling::span[1]').text()
                      #except :
                                 #do_m =''                                                       
                                 
                      #try:
                                 #opis1 = sess.at_xpath(u'//div[@class="extended-body"]').text()
                                 #opis = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", opis1)
                                 #opis = re.sub(u"[.,\-\s]{3,}", " ", opis)                                 
                      #except :
                                 #opis = ''
                      #try:
                                 #phone = sess.at_xpath('//div[@class="phone"]',timeout=1)['data-phone']
                      #except :
                                 #phone = ''
                      #try:
                                 #lico1 = sess.at_xpath(u'//div[@class="name"]').text()
                                 #lico = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", lico1)
                                 #lico = re.sub(u"[.,\-\s]{3,}", " ", lico)
                      #except :
                                 #lico = ''
                                 
                      #try:
                                 #comp = sess.at_xpath(u'//li[contains(text(),"Пассажирский лифт")]').text().replace(u' лифт','')
                      #except :
                                 #comp = ''
                                 
           
                      #try:
                                 #data = sess.at_xpath(u'//div[@class="row"]/div[contains(text(),"Данные ")]').text()
                                 #data1 = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", data)
                                 #data1 = re.sub(u"[.,\-\s]{3,}", " ", data1).replace(u'Данные обновлены ','').split(' ')[0].replace('-','.')
                      #except :
                                 #data1=''
                      #try:
                                 #mesto = sub+', '+punkt+', '+ray+', '+uliza
                      #except :
                                 #mesto =''
                      #try:
                                 #park = sess.at_xpath(u'//div[contains(text(),"Срок сдачи")]/following-sibling::div[1]').text()
                      #except :
                                 #park =''
                      #try:
                                 #vent = sess.at_xpath(u'//h2[@class="title--3rget"]').text()
                      #except :
                                 #vent =''
                      print('*'*50)
                      print ray 
                      #print punkt 
                      #print sub 
                      #print uliza
                      #print dom
                      #print seg
                      #print naz
                      #print price
                      #print klass
                      #print plosh
                      #print et
                      #print opis
                      #print phone
                      #print lico
                      #print cena_za
                      #print data1
                      #print mesto
                      
                      print('*'*50)
                      #ws.write(result, 0, sub)
                      ws.write(result, 1, ray)
                      #ws.write(result, 2, punkt)
                      #ws.write(result, 4, uliza)
                      #ws.write(result, 26, dom)
                      #ws.write(result, 6, seg)
                      #ws.write(result, 9, naz)
                      #ws.write(result, 11, price)
                      #ws.write(result, 28, cena_za)
                      #ws.write(result, 10, klass)
                      #ws.write(result, 14, plosh)
                      #ws.write(result, 15, et)
                      #ws.write(result, 33, et2)
                      #ws.write(result, 27, god)
                      ##ws.write(result, 14, zag)
                      #ws.write_string(result, 20, lin[z])
                      #ws.write(result, 30, data1)
                      #ws.write(result, 18, opis)
                      #ws.write(result, 21, phone)
                      #ws.write(result, 22, lico)
                      #ws.write(result, 19, u'БЦИнформ')
                      #ws.write(result, 31, datetime.today().strftime('%d.%m.%Y'))
                      #ws.write(result, 25, mesto)
                      #ws.write(result, 23, vent)
                      result+=1
                      time.sleep(1)
                      #sess.reset()
                      z=z+1
                      
except KeyboardInterrupt:
           pass
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')    




          


 


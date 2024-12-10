#!/usr/bin/python
# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,WebDriverException
from selenium.webdriver.common.by import By
import xlsxwriter
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import re
from datetime import datetime,timedelta


#profile =  webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/ljpce52l.default/') #Gui2
profile = webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/3missjz0.default/')#Gui1
#profile = webdriver.FirefoxProfile()
profile.native_events_enabled = False
#driver  = webdriver.Firefox(firefox_profile=profile,executable_path=r'/usr/local/bin/geckodriver',timeout=40)


#driver.set_window_position(0,0)
#driver.set_window_size(1000,720)


i = 38
ls= open('Links/zem.txt').read().splitlines()
dc = len(ls)

oper = u'Продажа'



while i < len(ls):
           print '********************************************************************************************'
           print i+1,'/',dc
           driver  = webdriver.Firefox(firefox_profile=profile,executable_path=r'/usr/local/bin/geckodriver',timeout=80)
           driver.set_window_position(0,0)
           driver.set_window_size(900,550)
           time.sleep(3)
           driver.get(ls[i])    
           print ls[i]
           time.sleep(2)
           sub = driver.find_element_by_xpath(u'//ul[@class="linklist navlinks"]/li[1]/a').text
           print sub
          
           workbook = xlsxwriter.Workbook(u'zem/Ners_zem'+str(i+1)+'.xlsx')
           ws = workbook.add_worksheet(u'Ners_Земля')
           ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
           ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
           ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
           ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
           ws.write(0, 4, u"УЛИЦА")
           ws.write(0, 5, u"ДОМ")
           ws.write(0, 6, u"ОРИЕНТИР")
           ws.write(0, 7, u"ТРАССА")
           ws.write(0, 8, u"УДАЛЕННОСТЬ")
           ws.write(0, 9, u"ОПЕРАЦИЯ")
           ws.write(0, 10, u"СТОИМОСТЬ")
           ws.write(0, 11, u"ЦЕНА_ЗА_СОТКУ")
           ws.write(0, 12, u"ПЛОЩАДЬ")
           ws.write(0, 13, u"КАТЕГОРИЯ_ЗЕМЛИ")
           ws.write(0, 14, u"ВИД_РАЗРЕШЕННОГО_ИСПОЛЬЗОВАНИЯ")
           ws.write(0, 15, u"ГАЗОСНАБЖЕНИЕ")
           ws.write(0, 16, u"ВОДОСНАБЖЕНИЕ")
           ws.write(0, 17, u"КАНАЛИЗАЦИЯ")
           ws.write(0, 18, u"ЭЛЕКТРОСНАБЖЕНИЕ")
           ws.write(0, 19, u"ТЕПЛОСНАБЖЕНИЕ")
           ws.write(0, 20, u"ОХРАНА")
           ws.write(0, 21, u"ДОПОЛНИТЕЛЬНЫЕ_ПОСТРОЙКИ")
           ws.write(0, 22, u"ОПИСАНИЕ")
           ws.write(0, 23, u"ИСТОЧНИК_ИНФОРМАЦИИ")
           ws.write(0, 24, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
           ws.write(0, 25, u"ТЕЛЕФОН")
           ws.write(0, 26, u"КОНТАКТНОЕ_ЛИЦО")
           ws.write(0, 27, u"КОМПАНИЯ")
           ws.write(0, 28, u"ДАТА_РАЗМЕЩЕНИЯ")
           ws.write(0, 29, u"ДАТА_ОБНОВЛЕНИЯ")
           ws.write(0, 30, u"ДАТА_ПАРСИНГА")
           ws.write(0, 31, u"ВИД_ПРАВА")
           ws.write(0, 32, u"МЕСТОПОЛОЖЕНИЕ")
           result= 1
           conv = [(u' августа ',u'.08.'), (u' июля ',u'.07.'),
                      (u' мая ',u'.05.'),(u' июня ',u'.06.'),
                      (u' марта ',u'.03.'),(u' апреля ',u'.04.'),
                      (u' января ',u'.01.'),(u' декабря ',u'.12.'),
                      (u' сентября ',u'.09.'),(u' ноября ',u'.11.'),
                      (u' февраля ',u'.02.'),(u' октября ',u'.10.'),
                      (u'сегодня,', (datetime.today().strftime('%d.%m.%Y'))),
                      (u'вчера','{:%d.%m.%Y}'.format(datetime.today() - timedelta(days=1)))]           
           lin = []





           while True:
                      print '********************',len(lin),'**********************'
                      try:
                                 #try:
                                            #WebDriverWait(driver,30).until(EC.element_to_be_clickable((By.XPATH,u'//div[@class="nolink"]/following::a[1]/div')))
                                            #print "Page is ready!"
                                 #except TimeoutException:
                                            #break
                                 time.sleep(1)
                                 for link in driver.find_elements_by_xpath(u'//h2/a[contains(@href,"object")]'):
                                            url = link.get_attribute('href')   
                                            print url
                                            lin.append(url)
                                 time.sleep(1)
                                 driver.execute_script("arguments[0].scrollIntoView(false);",driver.find_element_by_xpath(u'//div[@class="pagination"]'))
                                 time.sleep(1)
                                 driver.find_element_by_xpath(u'//div[@class="pagination"]/span/strong/following-sibling::a[1]').click()
                                 time.sleep(5)
                      except (NoSuchElementException,WebDriverException):
                                 z=0
                                 for line in lin:
                                            print z+1,'/',str(len(lin))+' - '+ sub+' '+str(i+1),'/',str(dc)     
                                            try:
                                                       driver.set_page_load_timeout(10)
                                                       driver.get(line)
                                            except TimeoutException:
                                                       driver.execute_script("window.stop();")
                                            except WebDriverException:
                                                       break
                                            print "Page is ready!"
                                            time.sleep(2)
                                            
                                            try:
                                                       ray = driver.find_element_by_xpath(u'//dt[contains(text(),"Район")]/following-sibling::dd').text                                                       
                                            except (NoSuchElementException,WebDriverException):
                                                       ray = ''
                                            try:
                                                       
                                                       if sub == u"Москва":
                                                                  punkt= u"Москва"
                                                       elif sub == u"Санкт-Петербург":
                                                                  punkt= u"Санкт-Петербург"
                                                       elif sub == u"Севастополь":
                                                                  punkt= u"Севастополь"
                                                       else:                                                       
                                                                  punkt= driver.find_element_by_xpath(u'//dt[contains(text(),"Населенный пункт:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       punkt = ''                                            
                                            try:
                                                       ter = driver.find_element_by_xpath(u'//dt[contains(text(),"Шоссе:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       ter =''
                                            try:
                                                       ul = driver.find_element_by_xpath(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text
                                                       if u' ул'in ul:
                                                                  uliza = ul.split(', ')[0]
                                                                  tip = ''
                                                       else:
                                                                  tip = ul.split(', ')[0]
                                                                  uliza = ''                                                       
                                            except (NoSuchElementException,WebDriverException):
                                                       uliza = ''
                                                       tip = ''
                                                       
                                            try:
                                                       dom = re.sub('[^\d]','',driver.find_element_by_xpath(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text.split(', ')[1])
                                            except (NoSuchElementException,IndexError,WebDriverException):
                                                       dom = ''                                            
                                           
                                            try:
                                                       price = driver.find_element_by_xpath(u'//dt[contains(text(),"Цена:")]/following-sibling::dd').text.split(' (')[0]
                                            except (NoSuchElementException,WebDriverException,IndexError):
                                                       price = ''
                                                       
                                            try:
                                                       cena_za = driver.find_element_by_xpath(u'//dl[@class="price"]/dd/span/span[contains(text(),"сотку")]').text
                                            except (NoSuchElementException,WebDriverException):
                                                       cena_za = ''
                                                       
                                            try: 
                                                       klass = driver.find_element_by_xpath(u'//span[contains(text(),"Использование")]/following::div[2]').text
                                            except (NoSuchElementException,WebDriverException):
                                                       klass =''
                                            try:
                                                       plosh = driver.find_element_by_xpath(u'//dt[contains(text(),"Площадь:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       plosh = ''
                                            try:
                                                       opis = driver.find_element_by_xpath(u'//div[@class="param"]/following-sibling::div[@class="info"]').text  
                                            except (NoSuchElementException,WebDriverException):
                                                       opis = ''
                                            
                                            try:
                                                       lico = driver.find_element_by_xpath(u'//div[@class="notes_contact"][1]').text
                                            except (NoSuchElementException,IndexError,WebDriverException):
                                                       lico = ''
                                                       
                                            try:
                                                       comp = driver.find_element_by_xpath(u'//a[@class="firm_link"]').text
                                            except (NoSuchElementException,WebDriverException):
                                                       comp = ''
                                            try:
                                                       vid_prava = driver.find_element_by_xpath(u'//p[@class="pbig_gray"][contains(text(),"Вид собственности")]/following::p[1]').text
                                            except (NoSuchElementException,WebDriverException):
                                                       vid_prava = ''
                                                       
                                            try:
                                                       vid_iz = driver.find_element_by_xpath(u'//p[@class="pbig_gray"][contains(text(),"Назначение земли")]/following::p[1]').text
                                            except (NoSuchElementException,WebDriverException):
                                                       vid_iz = ''                                                       
                                                       
                                            try:
                                                       d = driver.find_element_by_xpath(u'//div[contains(text(),"Дата размещения:")]').text.replace(u'Дата размещения: ','') 
                                                       data1 = reduce(lambda d, r: d.replace(r[0], r[1]), conv, d)
                                            except (NoSuchElementException,WebDriverException):
                                                       data1 = ''
                                            try:
                                                       dt = driver.find_element_by_xpath(u'//div[contains(text(),"Дата обновления:")]').text.replace(u'Дата обновления: ','')
                                                       data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)[:10]
                                            except (NoSuchElementException,WebDriverException):
                                                       data=''
                                            try:
                                                       mesto = driver.find_element_by_xpath(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       mesto =''
                                                       
                                            try:
                                                       time.sleep(1)
                                                       driver.execute_script("arguments[0].scrollIntoView(false);",driver.find_element_by_xpath(u'//div[@id="notes_contacts"]'))
                                                       time.sleep(1)
                                                       driver.find_element_by_xpath(u'//div[@id="get_phone"]').click()
                                                       time.sleep(2)                                                       
                                                       phone = re.sub('[^\d\+\,\:]', u'',driver.find_element_by_xpath(u'//div[@id="contact_phone"]/a').text)
                                            except (NoSuchElementException,WebDriverException):
                                                       phone = ''
                                                       
                                            print('*'*50)
                                            print ray 
                                            print punkt 
                                            print ter 
                                            print uliza
                                            print cena_za
                                            print price
                                            print klass
                                            print plosh
                                            print opis
                                            print phone
                                            print lico
                                            print comp
                                            print data
                                            print data1
                                            print mesto
                                            print('*'*50)
                                            ws.write(result, 0, sub)
                                            ws.write(result, 1, ray)
                                            ws.write(result, 2, punkt)
                                            ws.write(result, 3, tip)
                                            ws.write(result, 5, dom)
                                            ws.write(result, 7, ter)
                                            ws.write(result, 4, uliza)
                                            ws.write(result, 11, cena_za)
                                            ws.write(result, 14, klass)
                                            ws.write(result, 10, price)
                                            ws.write(result, 12, plosh)
                                            ws.write(result, 22, opis)
                                            ws.write(result, 23, u'Национальная единая риэлторская сеть')
                                            ws.write_string(result, 24, line)                                            
                                            ws.write(result, 25, phone)
                                            ws.write(result, 26, lico)
                                            ws.write(result, 27, comp)
                                            ws.write(result, 29, data)
                                            ws.write(result, 28, data1)
                                            ws.write(result, 9, oper)
                                            ws.write(result, 31, vid_prava)
                                            ws.write(result, 14, vid_iz)
                                            ws.write(result, 30, datetime.today().strftime('%d.%m.%Y'))
                                            ws.write(result, 32, mesto)
                                            result+=1
                                            z=z+1
                                 workbook.close()
                                 driver.close()
                                 time.sleep(5)                                 
                                 break
           i=i+1 



 


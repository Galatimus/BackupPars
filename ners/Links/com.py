#!/usr/bin/python
# -*- coding: utf-8 -*-

from selenium import webdriver
#from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import NoSuchElementException,WebDriverException,TimeoutException
from selenium.webdriver.common.by import By
import xlsxwriter
import time
import re
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')



profile =  webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/n1ddsdpx.default/') #Gui2
#profile = webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/3missjz0.default/')#Gui1
#profile = webdriver.FirefoxProfile()
profile.native_events_enabled = False
#driver  = webdriver.Firefox(firefox_profile=profile,executable_path=r'/usr/local/bin/geckodriver',timeout=40)


#driver.set_window_position(0,0)
#driver.set_window_size(1000,410)
#time.sleep(3)


i = 0
#ls= open('Links/com_arenda.txt').read().splitlines()
ls= open('Links/com_prod.txt').read().splitlines()
dc = len(ls)

#oper = u'Аренда'
oper = u'Продажа'



while i < len(ls):
           print '***********************'
           print i+1,'/',dc 
           driver  = webdriver.Firefox(firefox_profile=profile,executable_path=r'/usr/local/bin/geckodriver',timeout=600)
           driver.set_window_position(0,0)
           driver.set_window_size(900,600)
           time.sleep(3)
           driver.get(ls[i]) 
           print ls[i]
           time.sleep(2)
           #driver.execute_script("location.reload()")
           sub = driver.find_element_by_xpath(u'//ul[@class="linklist navlinks"]/li[1]/a').text
           print sub
          
           workbook = xlsxwriter.Workbook('com/Ners_com_'+oper+'_'+str(i+1)+'.xlsx')
           ws = workbook.add_worksheet(u'Ners_Коммерческая')
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
                                 time.sleep(2)
                      except (NoSuchElementException,WebDriverException):
                                 z=0
                                 for line in lin:
                                            print z+1,'/',str(len(lin))+' - '+ sub+' '+str(i+1),'/',str(dc)
                                            try:
                                                       #driver.set_page_load_timeout(30)                                                       
                                                       driver.get(line)
                                                       time.sleep(2)
                                                       print line
                                            except WebDriverException:
                                                       #driver.execute_script("window.stop();")
                                                       #time.sleep(5)
                                                       #driver.execute_script("location.reload()")
                                                       #time.sleep(5)
                                                       #continue
                                                       break
                                            
                                                       
                                            try:
                                                       ray = driver.find_element_by_xpath(u'//dt[contains(text(),"Район области:")]/following-sibling::dd').text                                                       
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
                                                       ter = driver.find_element_by_xpath(u'//dt[contains(text(),"Район:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       ter =''
                                            try:
                                                       uliza = driver.find_element_by_xpath(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       uliza = ''
                                            try:
                                                       naz = driver.find_element_by_xpath(u'//dt[contains(text(),"Тип объекта:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       naz = ''                                                       
                                            try:
                                                       price = driver.find_element_by_xpath(u'//dt[contains(text(),"Цена:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       price = ''
                                                       
                                            try:
                                                       cena_za = driver.find_element_by_xpath(u'//dt[contains(text(),"Этаж:")]/following-sibling::dd').text 
                                            except (NoSuchElementException,WebDriverException):
                                                       cena_za = ''
                                            try: 
                                                       klass = driver.find_element_by_xpath(u'//dt[contains(text(),"Этажность:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       klass =''
                                            try:
                                                       plosh = driver.find_element_by_xpath(u'//dt[contains(text(),"Общая площадь:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       plosh = ''
                                            try:
                                                       opis = driver.find_element_by_xpath(u'//div[@class="param"]/following-sibling::div[@class="info"]').text  
                                            except (NoSuchElementException,WebDriverException):
                                                       opis = ''
                                            
                                            try:
                                                       lico = driver.find_element_by_xpath(u'//a[@class="profile_link"]').text
                                            except (NoSuchElementException,WebDriverException,IndexError):
                                                       lico = ''
                                                       
                                            try:
                                                       comp = driver.find_element_by_xpath(u'//a[@class="firm_link"]').text
                                            except (NoSuchElementException,WebDriverException):
                                                       comp = ''
                                                       
                                            try:
                                                       dt = driver.find_element_by_xpath(u'//div[contains(text(),"Дата обновления:")]').text.replace(u'Дата обновления: ','')
                                                       data = reduce(lambda dt, r: dt.replace(r[0], r[1]), conv, dt)[:10]
                                            except (NoSuchElementException,WebDriverException):
                                                       data = ''
                                            try:
                                                       d = driver.find_element_by_xpath(u'//div[contains(text(),"Дата размещения:")]').text.replace(u'Дата размещения: ','')
                                                       data1 = reduce(lambda d, r: d.replace(r[0], r[1]), conv, d)
                                            except (NoSuchElementException,WebDriverException):
                                                       data1=''
                                            try:
                                                       mesto = driver.find_element_by_xpath(u'//dt[contains(text(),"Адрес:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       mesto =''
                                                       
                                            try:
                                                       metro = driver.find_element_by_xpath(u'//dt[contains(text(),"Метро:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       metro = ''
                                                       
                                            try:
                                                       dmetro = driver.find_element_by_xpath(u'//dt[contains(text(),"До метро:")]/following-sibling::dd').text
                                            except (NoSuchElementException,WebDriverException):
                                                       dmetro = ''
                                                       
                                            try:
                                                       zag = driver.find_element_by_xpath(u'//div[@class="notes_ntdt"]/following-sibling::h1').text
                                            except (NoSuchElementException,WebDriverException):
                                                       zag = ''
                                                       
                                            try:
                                                       lat = re.findall(u'notes_coord = (.*?)]',driver.page_source)[0].split(', ')[0][1:]
                                            except (NoSuchElementException,IndexError,WebDriverException):
                                                       lat = ''
                                                       
                                            try:
                                                       lng = re.findall(u'notes_coord = (.*?)]',driver.page_source)[0].split(', ')[1]
                                            except (NoSuchElementException,IndexError,WebDriverException):
                                                       lng = ''                                                       
                                                       
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
                                            print phone
                                            print lat
                                            print lng
                                            print('*'*50)
                                            ws.write(result, 0, sub)
                                            ws.write(result, 1, ray)
                                            ws.write(result, 2, punkt)
                                            ws.write(result, 3, ter)
                                            ws.write(result, 4, uliza)
                                            ws.write(result, 9, naz)
                                            ws.write(result, 15, cena_za)
                                            ws.write(result, 16, klass)
                                            ws.write(result, 11, price)
                                            ws.write(result, 14, plosh)
                                            ws.write(result, 18, opis)
                                            ws.write(result, 19, u'Национальная единая риэлторская сеть')
                                            ws.write_string(result, 20, line)                                            
                                            ws.write(result, 21, phone)
                                            ws.write(result, 22, lico)
                                            ws.write(result, 23, comp)
                                            ws.write(result, 30, data)
                                            ws.write(result, 29, data1)
                                            ws.write(result, 28, oper)
                                            ws.write(result, 31, datetime.today().strftime('%d.%m.%Y'))
                                            ws.write(result, 24, mesto)                                            
                                            ws.write(result, 26, metro)
                                            ws.write(result, 27, dmetro)
                                            ws.write(result, 33, zag)
                                            ws.write(result, 34, lat)
                                            ws.write(result, 35, lng)
                                            result+=1
                                            driver.delete_cookie
                                            z=z+1
                                 workbook.close()
                                 driver.close()
                                 time.sleep(3)                                 
                                 break
           i=i+1 



 


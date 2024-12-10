#!/usr/bin/python
# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import re
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os



#profile =  webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/46stx7t7.default/') #Gui1
profile = webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/yaun5l28.default/')#Gui2
#profile = webdriver.FirefoxProfile()
profile.set_preference('permissions.default.stylesheet', 2)
profile.set_preference('permissions.default.image', 2)
profile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', False)
profile.set_preference("javascript.enabled", False)
profile.native_events_enabled = False
options = webdriver.FirefoxOptions()
options.add_argument('-headless')
driver  = webdriver.Firefox(firefox_profile=profile,firefox_options=options,service_log_path=None,executable_path=r'/usr/local/bin/geckodriver',timeout=90)


#******************************************************************************************************
driver.get('https://torgi.gov.ru/lotSearch1.html?bidKindId=1')
time.sleep(2) 
driver.find_element_by_xpath(u'//label[contains(text(),"Тип имущества:")]/following::img[@class="ss_sprite ss_book_open"][1]').click()
time.sleep(2) 
driver.find_element_by_xpath(u'//span[@class="bold"][contains(text(),"Недвижимое имущество")]/..//input').click()
time.sleep(2) 
driver.find_element_by_xpath(u'//ins[contains(text(),"Выбрать")]').click()
time.sleep(2) 
select=Select(driver.find_element_by_name('common:country')).select_by_visible_text(u"РОССИЯ")
time.sleep(2) 

#raw_input('Введите число') 

driver.find_element_by_id(u'lot_search').click()
time.sleep(10)
driver.set_window_position(0,0)
driver.set_window_size(800,385)
time.sleep(5)

#Archive

#time.sleep(2)
#driver.get('http://torgi.gov.ru/lotSearchArchive.html')
#time.sleep(2)
#driver.find_element_by_xpath(u'//label[contains(text(),"Тип имущества:")]/following::img[@class="ss_sprite ss_book_open"][1]').click()
#time.sleep(1) 
#driver.find_element_by_xpath(u'//span[@class="bold"][contains(text(),"Недвижимое имущество")]/..//input').click()
#time.sleep(1) 
#driver.find_element_by_xpath(u'//ins[contains(text(),"Выбрать")]').click()
#time.sleep(1) 
#select=Select(driver.find_element_by_name('extended:country')).select_by_visible_text(u"РОССИЯ")
#driver.find_element_by_id(u'lot_search').click()
#time.sleep(10)
#*******************************************************************************************************
#driver.set_window_position(0,0)
#driver.set_window_size(1280,500)


dom = re.sub('[^\d]','',driver.find_element_by_xpath(u'//h2/span[2]').text)
print dom
lin = []
while True:
    try:
        WebDriverWait(driver,2000).until(EC.presence_of_element_located((By.XPATH,'//div[@id="over"][contains(@style,"none")]')))
        print "Page is ready!"
        time.sleep(1)
        for link in driver.find_elements_by_xpath(u'//a[contains(@title,"Просмотр")]'):
            url = link.get_attribute('href')   
            print url
            lin.append(url)
        time.sleep(1)
        #driver.execute_script("arguments[0].scrollIntoView(false);",driver.find_element_by_xpath(u'//tr[@class="navigation"]/td/div'))
        driver.execute_script("window.scrollTo(800, document.body.scrollHeight-520);")
        time.sleep(1) 
        #driver.find_element_by_xpath(u'//tr[@class="navigation"]/td/div/span/span/following::span[1]/a').click()
        driver.find_element_by_xpath(u'//tr[@class="navigation"]/td/div/a[contains(@title,"Перейти на одну страницу вперед")]').click()
    except NoSuchElementException:        
        print('Wait 2 sec...')
        time.sleep(2)
        print('Save it...')    
        command = 'mount -a'# cifs //192.168.1.6/e /home/oleg/Pars -o username=oleg,password=1122,iocharset=utf8,file_mode=0777,dir_mode=0777' #'mount -a -O _netdev'
        ##command = 'apt autoremove'
        p = os.system('echo %s|sudo -S %s' % ('1122', command))
        print p
        time.sleep(5)
        
        links = open('links/Torgi_Com_arenda.txt', 'w')
        for item in lin:
            links.write("%s\n" % item)
        links.close()
        driver.close()
        break
    except TimeoutException:
        print "Loading took too much time!"
        time.sleep(2)
        continue    
    print '**',len(lin),'**',str(dom)
print('Done!')






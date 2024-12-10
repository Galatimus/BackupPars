#!/usr/bin/python
# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import os
import re
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import random
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

profile = webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/yaun5l28.default/')#Gui2
profile.set_preference('permissions.default.stylesheet', 2)
profile.set_preference('permissions.default.image', 2)
profile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', False)
profile.set_preference("javascript.enabled", False)
profile.native_events_enabled = False
options = webdriver.FirefoxOptions()
options.add_argument('-headless')
driver  = webdriver.Firefox(firefox_profile=profile,firefox_options=options,executable_path=r'/usr/local/bin/geckodriver',service_log_path=None,timeout=90)

driver.set_window_position(0,0)
driver.set_window_size(800,750)
time.sleep(2) 
driver.get('https://torgi.gov.ru/lotSearch1.html?bidKindId=2')
time.sleep(5) 
select=Select(driver.find_element_by_name('common:country')).select_by_visible_text(u'РОССИЯ')
time.sleep(2) 
driver.find_element_by_id(u'lot_search').click()
time.sleep(10)
dom = re.sub('[^\d]','',driver.find_element_by_xpath(u'//h2/span[2]').text)
print dom

lin = []
while True:
    try:
        WebDriverWait(driver,10000).until(EC.presence_of_element_located((By.XPATH,'//div[@id="over"][contains(@style,"none")]')))
        #WebDriverWait(driver,60).until(EC.element_to_be_clickable((By.XPATH,u'//tr[@class="navigation"]/td/div/a[contains(@title,"Перейти на одну страницу вперед")]')))
        print "Page is ready!"
        time.sleep(1)
        for link in driver.find_elements_by_xpath(u'//a[contains(@title,"Просмотр")]'):
            url = link.get_attribute('href')   
            print url
            lin.append(url)
        time.sleep(1)
        #driver.execute_script("arguments[0].scrollIntoView(false);",driver.find_element_by_xpath(u'//tr[@class="navigation"]/td/div'))
        driver.execute_script("window.scrollTo(800, document.body.scrollHeight);")
        time.sleep(1)        
        driver.find_element_by_xpath(u'//tr[@class="navigation"]/td/div/a[contains(@title,"Перейти на одну страницу вперед")]').click()
        # //tr[@class="navigation"]/td/div/span/span/following::span[1]/a
    except NoSuchElementException:
        links = open('links/Torgi_Zem.txt', 'w')
        for item in lin:
            links.write("%s\n" % item)
        links.close()
        time.sleep(1) 
        driver.delete_all_cookies()
        driver.close()
        break
    except TimeoutException:
        print "Loading took too much time!"
        time.sleep(2)
        continue    
    print '**',len(lin),'**',str(dom)
print('Done!')




#https://torgi.gov.ru/?wicket:interface=:0:search_panel:resultTable:list:bottomToolbars:2:toolbar:span:navigator:first::IBehaviorListener:0:
#https://torgi.gov.ru/?wicket:interface=:0:search_panel:resultTable:list:bottomToolbars:2:toolbar:span:navigator:next::IBehaviorListener:0:-1


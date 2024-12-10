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
import sys
reload(sys)
sys.setdefaultencoding('utf-8')



#profile =  webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/3missjz0.default/') #Gui1
profile = webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/n1ddsdpx.default/')#Gui2
#profile = webdriver.FirefoxProfile()
profile.native_events_enabled = False
driver  = webdriver.Firefox(firefox_profile=profile,executable_path=r'/usr/local/bin/geckodriver',timeout=600)
driver.set_window_position(0,0)
driver.set_window_size(900,500)




i = 0
ls= open('Links/zem.txt').read().splitlines()
dc = len(ls)





while i < len(ls):
           print '***********'
           print i+1,'/',dc  
           #driver  = webdriver.Firefox(firefox_profile=profile,executable_path=r'/usr/local/bin/geckodriver',timeout=40)
           #driver.set_window_position(0,0)
           #driver.set_window_size(900,600)
           time.sleep(2)           
           driver.get(ls[i])    
           print ls[i]
           time.sleep(3)
           lin = []
           while True:
                      
                      try:
                                 WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//select[@id="sf_change"]/following-sibling::span[contains(text(),"Сортировка: новые с начала")]')))
                                 print "Page is ready!"
                                 time.sleep(1)
                                 for link in driver.find_elements_by_xpath(u'//h2/a[contains(@href,"object")]'):
                                            url = link.get_attribute('href')   
                                            print url
                                            lin.append(url)
                                 time.sleep(1)
                                 print '***',len(lin),'***'
                                 print i+1,'/',dc
                                 driver.execute_script("arguments[0].scrollIntoView(false);",driver.find_element_by_xpath(u'//div[@class="pagination"]'))                                 
                                 time.sleep(1)
                                 driver.find_element_by_xpath(u'//div[@class="pagination"]/span/strong/following-sibling::a[1]').click()
                      except (NoSuchElementException,WebDriverException):
                                 links = open('ners_zem.txt', 'a')
                                 for item in lin:
                                            links.write("%s\n" % item)
                                 links.close()
                                 time.sleep(1) 
                                 driver.get("about:blank")
                                 driver.delete_all_cookies()
                                 time.sleep(2)                                 
                                 break                                 
                      except TimeoutException:
                                 print "Loading took too much time!"
                                 time.sleep(2)
                                 continue                                 
           
           i=i+1      
           
driver.close()


 


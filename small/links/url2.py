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



profile =  webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/n1ddsdpx.default/') #Gui2
#profile = webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/3missjz0.default/')#Gui1
#profile = webdriver.FirefoxProfile()
profile.native_events_enabled = False
driver  = webdriver.Firefox(firefox_profile=profile,executable_path=r'/usr/local/bin/geckodriver',timeout=600)




driver.get("https://dom.sakh.com/land/")


driver.set_window_position(0,0)
driver.set_window_size(800,600)

time.sleep(3) 





lin = []
while True:
           try:
                      #WebDriverWait(driver,2000).until(EC.presence_of_element_located((By.XPATH,'//div[@id="over"][contains(@style,"none")]')))
                      #WebDriverWait(driver,60).until(EC.element_to_be_clickable((By.XPATH,u'//tr[@class="navigation"]/td/div/a[contains(@title,"Перейти на одну страницу вперед")]')))
                      print "Page is ready!"
                      time.sleep(1)
                      for link in driver.find_elements_by_xpath(u'//div[@class="actions noprint"]/following-sibling::a'):
                                 url = link.get_attribute('href')   
                                 print url
                                 lin.append(url)
                      time.sleep(1)
                      driver.execute_script("arguments[0].scrollIntoView(false);",driver.find_element_by_xpath(u'//div[@class="breadcrumbs bound"]'))                                 
                      time.sleep(1)        
                      driver.find_element_by_xpath(u'//a[@class="step"][contains(text(),"следующая")]').click()
                      time.sleep(5)
           except NoSuchElementException:
                      links = open('zem.txt', 'w')
                      for item in lin:
                                 links.write("%s\n" % item)
                      links.close()
                      driver.close()
                      break
           #except TimeoutException:
                      #print "Loading took too much time!"
                      #time.sleep(2)
                      #continue    
           print '********************',len(lin),'**********************'

print'DONE!!!!'



 


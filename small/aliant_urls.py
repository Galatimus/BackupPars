#!/usr/bin/python
# -*- coding: utf-8 -*-




from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,ElementNotVisibleException,ElementNotInteractableException
from selenium.webdriver.common.by import By
import time
import random
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import TimeoutException






profile = webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/yaun5l28.default/')#Gui2
profile.set_preference('permissions.default.stylesheet', 2)
profile.set_preference('permissions.default.image', 2)
profile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', False)
profile.set_preference("javascript.enabled", False)
profile.native_events_enabled = False
driver  = webdriver.Firefox(firefox_profile=profile,executable_path=r'/usr/local/bin/geckodriver',service_log_path=None,timeout=90)
driver.set_window_position(0,0)
driver.set_window_size(800,750)


driver.get("http://aliant.pro/catalog/commercial/")

time.sleep(5)

while True:
       try:
              driver.execute_script("arguments[0].scrollIntoView(false);",driver.find_element_by_xpath(u'//div[@class="contactUsLeft"]'))
              #WebDriverWait(driver,60).until(EC.presence_of_element_located((By.XPATH,'//div[@id="preloader"][contains(@style,"none")]')))
              print "Page is ready!"
              time.sleep(1)
              driver.find_element_by_xpath(u'//button[@class="waves-effect waves-light btn btn-more"]').click()
              time.sleep(1)
              print('Done!') 
       except (ElementNotVisibleException,ElementNotInteractableException):
              for link in driver.find_elements_by_xpath(u'//a[@class="item"]'):
                     url = link.get_attribute('href')   
                     print url
                     li = open('aliant.txt', 'a')
                     li.write(url + '\n')
                     li.close()                     
              driver.close()
              print('Done All') 
              break
              
              
              
       time.sleep(2)       
                   
 
    
   

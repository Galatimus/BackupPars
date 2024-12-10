#!/usr/bin/python
# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,WebDriverException
from selenium.webdriver.common.by import By
import xlsxwriter
import os
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import re
from datetime import datetime,timedelta
import sys
reload(sys)
sys.setdefaultencoding('utf-8')





profile =  webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/cqryyjra.default/') #Gui1
#profile = webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/o5wsi6o1.default/')#Gui2
profile.set_preference('permissions.default.stylesheet', 2)
profile.set_preference('permissions.default.image', 2)
profile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', False)
profile.set_preference("javascript.enabled", False)
profile.native_events_enabled = False
driver  = webdriver.Firefox(firefox_profile=profile,capabilities={"marionette": False},timeout=90)



#ua = dict(DesiredCapabilities.PHANTOMJS)
#ua["phantomjs.page.settings.userAgent"] = ("Mozilla/5.0 (Windows NT 6.1; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0")
#service_args = ['--proxy='+proxy,'--proxy-type=http',]
##service_args=['--ignore-ssl-errors=true', '--ssl-protocol=any']
#driver = webdriver.PhantomJS(service_args=service_args)


driver.set_window_position(0,0)
driver.set_window_size(800,500)




i = 0
ls= open('Links/com_all.txt').read().splitlines()
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
                                 #WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//title')))
                                 print "Page is ready!"
                                 time.sleep(3)
                                 for link in driver.find_elements_by_xpath(u'//div[@class="media-body"]/div/following-sibling::a'):
                                            url = link.get_attribute('href')   
                                            print url
                                            lin.append(url)
                                 #time.sleep(1)
                                 print '***',len(lin),'***'
                                 print i+1,'/',dc
                                 driver.execute_script("arguments[0].scrollIntoView(false);",driver.find_element_by_xpath(u'//div[@class="pagination"]'))                                 
                                 time.sleep(1)
                                 driver.find_element_by_xpath(u'//div[@class="pagination"]/span/strong/following-sibling::a[1]').click()
                                 time.sleep(3)
                      except (NoSuchElementException,WebDriverException):
                                 links = open('ners_com.txt', 'a')
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
                                 time.sleep(1) 
                                 driver.get("about:blank")
                                 driver.delete_all_cookies()                                 
                                 time.sleep(2)
                                 continue                                 
           
           i=i+1      
           
driver.close()
print('Done!') 
time.sleep(5)
os.system("/home/oleg/pars/ners/comm.py")


 


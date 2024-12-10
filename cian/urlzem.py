#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
import os
import math
import time
from datetime import datetime,timedelta
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab


logging.basicConfig(level=logging.DEBUG)

try:
    os.remove('/home/oleg/pars/cian/cian_zem.txt')
    print 'Удаляем: '
except (IOError, OSError):
    print 'Нет файла'


i = 0
ls= open('Links/zemm.txt').read().splitlines()
dc = len(ls)

places = []

while i < len(ls):
    print '********************************************',i+1,'/',dc,'*******************************************'
    page = ls[i]
    lin = []    
    class Brsn_Com(Spider):
	
	
	
	def prepare(self):
	    self.f = page
	def task_generator(self):
	    for x in range(1,60):
		link = self.f+'&p='+str(x)
		yield Task ('post',url=link.replace(u'&p=1',''),refresh_cache=True,network_try_count=20)
    
	def task_post(self,grab,task):
	    for elem in grab.doc.select(u'//a[contains(@href,"sale/suburban")]'):
                url = grab.make_url_absolute(elem.attr('href')) 
                lin.append(url)
	    for elem in grab.doc.select(u'//a[contains(@href,"rent/suburban")]'):
	        url = grab.make_url_absolute(elem.attr('href'))   
	        lin.append(url) 
		
		
	    print '***',len(lin),'****'
	    print '***',i+1,'/',dc,'****'
			
    
    bot = Brsn_Com(thread_number=5,network_try_limit=200)
    bot.load_proxylist('../tipa.txt','text_file')
    bot.create_grab_instance(timeout=500, connect_timeout=500)
    bot.run()
    print('Done!')
    time.sleep(1) 
    lin = list(set(lin))
    print 'Save...' 
    print '***',len(lin),'****'
    time.sleep(2) 
    for item in lin:
        places.append(item)
    print 'Total...',len(places)
    time.sleep(1)    
    i=i+1 
liks = open('cian_zem.txt', 'w')
for x in range(3):
    for itm in places:
        liks.write("%s\n" % itm)
liks.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/cian/zem.py")

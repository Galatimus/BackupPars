#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
import re
import os
import math
import time
from datetime import datetime,timedelta
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab


logging.basicConfig(level=logging.DEBUG)




i = 0
ls= open('Links/flats.txt').read().splitlines()
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
		yield Task ('post',url=self.f+'&p='+str(x),refresh_cache=True,network_try_count=20)
            yield Task ('post',url=self.f,refresh_cache=True,network_try_count=20)


	def task_post(self,grab,task):
	    for elem in grab.doc.select(u'//a[contains(@href,"sale/flat")]'):
                url = grab.make_url_absolute(elem.attr('href')) 
                lin.append(url)
	    for elem in grab.doc.select(u'//a[contains(@href,"rent/flat")]'):
	        url = grab.make_url_absolute(elem.attr('href'))   
	        lin.append(url) 
		
	    print('*'*100)
            print 'Ready - '+str(len(lin))    
	    print '***',i+1,'/',dc,'***'
	    print('*'*100)
    
    bot = Brsn_Com(thread_number=5,network_try_limit=200)
    bot.load_proxylist('../tipa.txt','text_file')
    bot.create_grab_instance(timeout=50, connect_timeout=50)
    bot.run()
    print('Done!')
    time.sleep(1) 
    lin = list(set(lin))
    print 'Save...' 
    print '************************',len(lin),'**********************************'
    time.sleep(2)    
    for item in lin:
        places.append(item)
    print 'Total...',len(places)
    time.sleep(1)    
    i=i+1    
liks = open('cian_flats.txt', 'w')
for x in range(3):
    for itm in places:
        liks.write("%s\n" % itm)
liks.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/cian/flat.py")

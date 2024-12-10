#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import re
#from datetime import datetime
import math
import time
from datetime import datetime,timedelta
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
from grab import Grab


logging.basicConfig(level=logging.DEBUG)




i = 0
l1 = 2
ls= open('Links/new5.txt').read().splitlines()
dc = len(ls)
v = ['shoppingCenter','businessCenter','warehouse']
o = ['rent','sale']


#shoppingCenter
#businessCenter
#warehouse
while i < len(ls):
    print '********************************************',i+1,'/',dc,'*******************************************'
    page = ls[i]
    oper = 'sale'
    vid = v[l1]
    lin = []
    class Brsn_Com(Spider):
	
	
	
	def prepare(self):
	    self.f = page
	def task_generator(self):
	    for x in range(1,60):
		link = 'https://www.cian.ru/bs_centers/list/?bs_center_type='+vid+'&deal_type='+oper+'&engine_version=2&offer_type=offices'+'&p='+str(x)+self.f
		yield Task ('post',url=link,refresh_cache=True,network_try_count=50)
	def task_post(self,grab,task):
	    for elem in grab.doc.select(u'//h3/a'):
                url = grab.make_url_absolute(elem.attr('href')) 
                lin.append(url)
	    #for elem in grab.doc.select(u'//a[contains(@href,"rent/suburban")]'):
	        #url = grab.make_url_absolute(elem.attr('href'))   
	        #lin.append(url) 
		
		
	    
	    print '***',i+1,'/',dc,'****',len(lin),'****'
	    
    bot = Brsn_Com(thread_number=5,network_try_limit=500)
    bot.load_proxylist('../tipa.txt','text_file')
    bot.create_grab_instance(timeout=500, connect_timeout=500)
    bot.run()
    print('Done!')
    time.sleep(1) 
    lin = list(set(lin))
    print 'Save...' 
    print '***',len(lin),'****'
    time.sleep(2) 
    links = open('cian_bc.txt', 'a')
    for item in lin:
        links.write("%s\n" % item)
    links.close()
    time.sleep(1)    
    i=i+1 
   

    
    #page = ls[i]

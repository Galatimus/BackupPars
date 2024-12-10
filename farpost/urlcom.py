#! /usr/bin/env python
# -*- coding: utf-8 -*-


from grab.spider import Spider,Task
import logging
import time
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

try:
     os.remove('/home/oleg/pars/farpost/faprost_com.txt')
     print 'Удаляем: '
except (IOError, OSError):
     print 'Нет файла'



class Farpost_Com(Spider):
     def prepare(self):
	  self.lin = []
     def task_generator(self):
	  for x in range(1,67):#52
	       yield Task ('post',url='https://www.farpost.ru/realty/rent_business_realty/?page=%d'%x,refresh_cache=True,network_try_count=100)
	  for x1 in range(1,10):#9
	       yield Task ('post',url='https://www.farpost.ru/realty/rent_garage/?page=%d'%x1,refresh_cache=True,network_try_count=100)
	  for x2 in range(1,6):#4
	       yield Task ('post',url='https://www.farpost.ru/rest/hotels/?page=%d'%x2,refresh_cache=True,network_try_count=100)	 
	  for x3 in range(1,50):#51
	       yield Task ('post',url='https://www.farpost.ru/realty/sale_garage/?page=%d'%x3,refresh_cache=True,network_try_count=100)
	  for x4 in range(1,40):#30
	       yield Task ('post',url='https://www.farpost.ru/realty/sell_business_realty/?page=%d'%x4,refresh_cache=True,network_try_count=100)
	  
			      
     def task_post(self,grab,task): 
	  for elem in grab.doc.select(u'//td[@class="imageCell"]/a'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       self.lin.append(ur)
	  print('*'*50)
	  print len(self.lin)
	  print('*'*50)
        
     
bot = Farpost_Com(thread_number=1,network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=5000, connect_timeout=5000)
try:
     bot.run()
except KeyboardInterrupt:
     pass
print('Wait 2 sec...')
time.sleep(2)
print len(bot.lin)
links = open('faprost_com.txt', 'w')
for x in range(10):
     for item in bot.lin:
          links.write("%s\n" % item)
links.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/farpost/comm.py")
#os.system("/home/oleg/pars/small/roszem.py")









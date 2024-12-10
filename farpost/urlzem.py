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
     os.remove('/home/oleg/pars/farpost/faprost_zem.txt')
     print 'Удаляем: '
except (IOError, OSError):
     print 'Нет файла'



class Farpost_Com(Spider):
     def prepare(self):
	  self.lin = []
     def task_generator(self):
	  for x in range(93):#78
               yield Task ('post',url='https://www.farpost.ru/realty/land/?page=%d'%x,refresh_cache=True,network_try_count=100)
          for x1 in range(4):#4
	       yield Task ('post',url='https://www.farpost.ru/realty/land-rent/?page=%d'%x1,refresh_cache=True,network_try_count=100)

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
links = open('faprost_zem.txt', 'w')
for x in range(10):
     for item in bot.lin:
          links.write("%s\n" % item)
links.close()
print('Done')
time.sleep(5)
os.system("/home/oleg/pars/farpost/zemm.py")
#os.system("/home/oleg/pars/small/roszem.py")









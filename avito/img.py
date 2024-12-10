#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
import logging
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import re
import random
from cStringIO import StringIO
import pytesseract
from PIL import Image 
import os
import time
import base64
import sys
reload(sys)
sys.setdefaultencoding('utf-8')



logging.basicConfig(level=logging.DEBUG)





class Avito(Spider):

     def prepare(self):
	  self.f = 'https://www.avito.ru/respublika_krym/kommercheskaya_nedvizhimost'
	  self.pag = 50
	  self.result= 1 

     def task_generator(self):
	  for x in range(1,self.pag+1):
	       yield Task ('post',url=self.f+'?p=%d'%x,refresh_cache=True,network_try_count=100)	  

	  
     def task_post(self,grab,task):
	  for elem in grab.doc.select(u'//a[@class="item-description-title-link"]'):
	       ur = grab.make_url_absolute(elem.attr('href'))  
	       #print ur	      
	       yield Task('item',url=ur,refresh_cache=True,network_try_count=100)
	     
     def task_item(self, grab, task):
	   
          ad_id= re.sub(u'[^\d]','',task.url[-10:])
	  #agent = random.choice(list(open('../agents.txt').read().splitlines()))
	  agent = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:65.0) Gecko/20100101 Firefox/65.0'
	  box = 'https://api.site-shot.com/?width=1024&height=0&zoom=100&scaled_width=1024&full_size=1&delay_time=2000&timeout=60000&format=PNG&user_agent='+agent+'&response_type=json&url='+task.url
	  #box = 'https://mini.s-shot.ru/1024x0/JPEG/1024/Z100/?'+task.url
	  yield Task('shot',url=box,ad_id=ad_id,refresh_cache=True,network_try_count=100)

     
	       
     def task_shot(self, grab, task):
	  try:
	       print('*'*50)
	       print task.ad_id
	       data_image64 = grab.doc.json['image'].split(',')[1] 
	       imgdata = base64.b64decode(data_image64)
	       im = Image.open(StringIO(imgdata))
	       path = 'img/Avito_%s.jpg' % task.ad_id
	       im.save(path)
	       print 'Screenshot OK'
	       print 'Ready - '+str(self.result)
               print 'Tasks - %s' % self.task_queue.size()
               print('*'*50)
               self.result+= 1
	       del im
	  except (IndexError,ValueError,GrabNetworkError,GrabTimeoutError,IOError):
	       pass
		       
	     

bot = Avito(thread_number=5, network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file',proxy_type='http')
bot.create_grab_instance(timeout=50, connect_timeout=50)
try:
     bot.run()
except KeyboardInterrupt:
     pass
print('Wait 2 sec...')
time.sleep(1)
print('Done!')

     

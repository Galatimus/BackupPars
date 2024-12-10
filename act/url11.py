#!/usr/bin/python
# -*- coding: utf-8 -*-





from grab.spider import Spider,Task
import grab.spider.queue_backend
import grab.spider.queue_backend.memory
import grab.transport
import grab.transport.curl
import logging
import time
import xlrd
import os
import re
from datetime import datetime
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


   
logging.basicConfig(level=logging.DEBUG)

       
name ='Книга'

class Gis(Spider):
       
       
       def prepare(self):
              self.rb = xlrd.open_workbook(name+'.xlsx',on_demand=True)
              self.sheet = self.rb.sheet_by_index(0)              
              self.workbook = xlsxwriter.Workbook(u'Актуальность11.xlsx')#+'True'+'.xlsx')
              self.ws = self.workbook.add_worksheet()
              self.ws.write(0,0, u"КодПредложения")
              self.ws.write(0,1, u"Источник")
              self.ws.write(0,2, u"Ссылка")
              self.ws.write(0,3, u"Актуальность")
              self.row= 1  

              
       def task_generator(self):
              for ak in range(1,self.sheet.nrows):
                     #time.sleep(1)
                     links = self.sheet.cell_value(ak,1).strip()
                     cod = self.sheet.cell_value(ak,0)
                     ist = ''#self.sheet.cell_value(ak,1).strip()
                     yield Task ('post',url= links,cod=cod,network_try_count=5)
        
                     
       def task_post(self,grab,task):
              #print task.url,task.ist
              
              #if 'Domofond' in opcs:
              if 'move.ru/' in task.url:
                     if grab.doc.select(u'//p[@class="block-user__show-telephone_number"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'        
       
       
              elif 'yuga.ru/' in task.url:
                     if grab.doc.select(u'//div[@itemprop="description"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
       
              elif 'dom.59.ru/' in task.url:
                     if grab.doc.select(u'//a[@class="all_ads_user_link"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'        
       
              #if '45_ru' in opcs:       
              elif 'dom.45.ru/' in task.url:
                     if grab.doc.select(u'//a[@class="all_ads_user_link"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
       
              #if '72_ru' in opcs:       
              elif 'dom.72.ru/' in task.url:
                     if grab.doc.select(u'//a[@class="all_ads_user_link"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              #if 'Domchel_ru' in opcs:        
              elif 'domchel.ru/' in task.url:
                     if grab.doc.select(u'//a[@class="all_ads_user_link"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              #if 'Farpost' in opcs: 
              elif 'farpost.ru/' in task.url:
                     if grab.doc.select(u'//strong[contains(text(),"Объявление находится в архиве и может быть неактуальным.")]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'
       
              #if 'Infoline' in opcs:
              elif 'vrx.ru/' in task.url:
                     if grab.doc.select(u'//td[contains(text(),"Операция:")]/following-sibling::td').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
       
              #if 'Tulahouse_ru' in opcs:
              elif 'tulahouse.ru/' in task.url:
                     if grab.doc.select(u'//p[@class="item-description"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              #if 'ГдеЭтотДом' in opcs:
              elif 'gdeetotdom.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="adv_status"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'
       
              #if 'Недвижимость_Астрахани' in opcs:
              elif 'n30' in task.url :
                     if grab.doc.select(u'//td[@class="thprice"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              #if 'Недвижимость_Екатеринбурга' in opcs:
              elif 'kvadrat66.ru/' in task.url:
                     if grab.doc.select(u'//td[@class="thprice"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              #if 'Недвижимость_Кемерово' in opcs:
              elif 'kemdom.ru/' in task.url:
                     if grab.doc.select(u'//td[@class="thprice"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              #if 'Недвижимость_Саратова' in opcs:
              elif 'kvadrat64.ru/' in task.url:
                     if grab.doc.select(u'//td[@class="thprice"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
       
              #if 'Недвижимость_и_цены' in opcs:
              elif 'dmir.ru/' in task.url:
                     if grab.doc.select(u'//span[@id="price_offer"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              #if 'Росриэлт_Недвижимость' in opcs:
              elif 'rosrealt.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="section_right"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              #if 'Уральская_палата_недвижимости' in opcs:
              elif 'upn.ru/' in task.url:
                     if grab.doc.select(u'//div[@id="ctl00_VOI_pnError"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'
       
              #if 'Циан' in opcs:
              elif 'cian.ru/' in task.url:
                     if grab.doc.select(u'//span[@class="object_descr_warning object_descr_warning_red"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'
       
              #if 'Avito' in opcs:
              elif 'avito.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="item-phone js-item-phone"]/div').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
       
              #if 'IRR' in opcs:
              elif 'irr.ru/' in task.url:
                     if grab.doc.select(u'//@data-phone').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'                    
       
       
              #if 'Mirkvartir' in opcs:
              elif 'mirkvartir.ru/' in task.url:
                     if grab.doc.select(u'//span[@class="phones"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'                     
       
              #if 'Theproperty' in opcs:
              elif 'theproperty.ru/' in task.url:
                     if grab.doc.select(u'//p[@class="archive"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'                     
       
              #if 'RealtyMag' in opcs:
              elif 'realtymag.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="error-page__code"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'                     
       
       
              #if 'Необходимая_недвижимость' in opcs:
              elif 'nndv.ru/' in task.url:
                     if grab.doc.select(u'//td[@class="paddLR5TB2"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'                     
       
              #if 'Mlsn' in opcs:
              elif 'mlsn.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="NotFound__base"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'                     
       
              elif 'life-realty.ru/' in task.url :
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'citystar.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'citystar74.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'realtyekaterinburg.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'                        
       
              elif 'n1.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'rosnedv.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ayax.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'qp.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'                            
       
              elif 'doska.ru/' in task.url:
                     if grab.doc.select(u'//span[@id="phone_td_1"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ners.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="notes_publish_status"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'
       
              elif 'ngs.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="card__phones-container"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'bn.ru/' in task.url:
                     if grab.doc.select(u'//dt[contains(text(),"Телефон")]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'nmls.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="mb10"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'                            
       
              elif 'bkn42.ru/' in task.url:
                     nek = grab.doc.select(u'//title').text().split(' ')[0]
                     if nek == u'ПРОДАНО':
                            akt = 'False'
                     elif nek ==u'СДАНО':
                            akt = 'False'
                     else:
                            akt = 'True'
       
              elif 'realtyvision.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'dom43.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'rk-region.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'    
              elif 'home29.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="message"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'
       
              ############################Nedvizhka.RU############################################################
       
              elif 'ned22.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False' 
       
              elif 'eest.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ned30.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ned31.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ned33.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'nedvizhka.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'vnk39.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'radver.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif '23estate.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ned77.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False' 
       
              elif 'ned74.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif '52metra.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'prmrealty.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ned02.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ned61.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'realt66.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'ned72.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
                            
              elif 'vse42.ru' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
              
              elif 'brsn.ru/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'mob_sellcom' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'mob_givecom' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
       
              elif 'advecs.com/' in task.url:
                     if grab.response.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'                     
              else:
                     akt =''
        
              self.ws.write(self.row, 3, akt)
              self.ws.write_string(self.row, 2, task.url)
              #self.ws.write(self.row, 1, task.ist)
              self.ws.write(self.row, 0, task.cod)
              print('*'*50)
              print akt
              print 'Ready - '+str(self.row)+'/'+str(self.sheet.nrows)
              print 'Tasks - %s' % self.task_queue.size()
              print name
              print('*'*50) 
              self.row+= 1              
                     
              #if self.row > 100:
                     #self.stop()                      
              
              
bot = Gis(thread_number=50, network_try_limit=50)
bot.load_proxylist('ftp://Oleg:walter2005@192.168.1.6/tipa.txt','url',proxy_type='http')
#bot.create_grab_instance(timeout=50, connect_timeout=1000)
try:
       bot.run()
except KeyboardInterrupt:
       pass
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')    
bot.workbook.close()
#workbook.close()
print('Done!')


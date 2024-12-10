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


name ='comm1'
time.sleep(2)
os.system('echo %s|sudo -S %s' % ('1122', 'service mongod start'))
time.sleep(2)
os.system('echo %s|sudo -S %s' % ('1122', 'mongo actualcom1 --eval "db.dropDatabase()"'))

class Actuall(Spider):


       def prepare(self):
              self.rb = xlrd.open_workbook(name+'.xlsx',on_demand=True)
              self.sheet = self.rb.sheet_by_index(0)
              self.workbook = xlsxwriter.Workbook(u'com/ActCom1.xlsx')#+datetime.today().strftime('%d.%m.%Y')+'.xlsx')
              self.ws = self.workbook.add_worksheet()
              self.ws.write(0,0, u"КодПредложения")
              self.ws.write(0,1, u"Источник")
              self.ws.write(0,2, u"Ссылка")
              self.ws.write(0,3, u"Актуальность")
              self.row= 1


       def task_generator(self):
              for ak in range(1,self.sheet.nrows):
                     #time.sleep(1)
                     links = self.sheet.cell_value(ak,2)
                     cod = self.sheet.cell_value(ak,0)
                     ist = self.sheet.cell_value(ak,1).lower()

                     #if not 'theproperty' in links:
                     yield Task ('post',url= links,ist=ist,cod=cod,network_try_count=2)

       def task_post(self,grab,task):
              #print task.url,task.ist

              if 'avito.ru' in task.ist:
                     if grab.doc.select(u'//div[@class="item-phone js-item-phone"]/div').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'bbport.ru/' in task.url:
                     if grab.doc.code ==200:
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
              elif 'domchel.ru' in task.ist:
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
                     if grab.doc.select(u'//div[@class="mrk bphone"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'


              #if 'Tulahouse_ru' in opcs:
              elif 'tulahouse.ru/' in task.url:
                     if grab.doc.select(u'//p[@class="item-description"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'kalugahouse.ru/' in task.url:
                     if grab.doc.select(u'//p[@class="item-description"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'vladimirhouse.ru/' in task.url:
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
              elif 'gde.ru' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              #if 'Росриэлт_Недвижимость' in opcs:
              elif 'rosrealt.ru/' in task.url:
                     if grab.doc.select(u'//p[@class="pbig_gray_contact"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'

              #if 'Уральская_палата_недвижимости' in opcs:
              elif 'upn.ru/' in task.url:
                     if grab.doc.select(u'//h1[contains(text(),"Запрашиваемая страница не найдена")]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'

              #if 'Циан' in opcs:
              elif 'циан' in task.ist:
                     if grab.doc.select(u'//div[contains(text(),"Объявление снято с публикации")]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'

              #if 'Avito' in opcs:
              elif 'ryazanhouse.ru' in task.ist:
                     if grab.doc.select(u'//p[@class="item-description"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'


              #if 'IRR' in opcs:
              elif 'из рук в руки' in task.ist:
                     if grab.doc.select(u'//div[@class="wrapperFullWidth"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'


              #if 'Mirkvartir' in opcs:
              elif 'mirkvartir' in task.url:
                     if grab.doc.select(u'//div[@class="l-object-description"]/p').exists() == True:
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
              elif 'realtymag' in task.url:
                     if grab.doc.code == 200:
                            akt = 'True'
                     else:
                            akt = 'False'


              #if 'Необходимая_недвижимость' in opcs:
              elif 'nndv.ru/' in task.url:
                     if grab.doc.select(u'//label[contains(text(),"Стоимость:")]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'

              #if 'Mlsn' in opcs:
              elif 'mlsn.ru' in task.ist:
                     if grab.doc.select(u'//div[contains(text(),"Объявление находится в архиве. Контакты продавца скрыты")]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'

              elif 'life-realty.ru/' in task.url :
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'ситистар' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'элиант Недвижимость - ан' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Любимый Город - АН' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'torgi.gov.ru/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'rosnedv.ru/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'ayax.ru/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'купи.ру' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'doska.ru/' in task.url:
                     if grab.doc.select(u'//span[@id="phone_td_1"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'национальная единая риэлторская сеть' in task.ist:
                     if grab.doc.select(u'//div[@id="contact_phone"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'нгс.недвижимость' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'бюллетень недвижимости' in task.ist:
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
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'dom43.ru/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'irk.ru/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'rk-region.ru/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'
              elif 'home29.ru/' in task.url:
                     if grab.doc.select(u'//div[@class="message"]').exists() == True:
                            akt = 'False'
                     else:
                            akt = 'True'

              elif 'move.ru' in task.ist:
                     if grab.doc.select(u'//p[@class="block-user__show-telephone_number"]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'
              ############################Nedvizhka.RU############################################################

              elif 'nedvizhka.ru' in task.ist:
                     if grab.doc.select(u'//h2[contains(text(),"Страница не найдена")]').exists() == True :#or grab.doc.code ==404:
                            akt = 'False'
                     else:
                            akt = 'True'

              elif 'Золотая Середина - АН' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Вестум.ru' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Домовой45' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Олимп - АН' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Градстрой - АН' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Дома-24' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Квартал - АН' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Издательский дом Ярмарка' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'Связист - АН' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'N1.ru' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'market.sakh.com' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'raui' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'youla' in task.url:
                     if u'Неактивно' in grab.doc.select(u'//title').text():
                            akt = 'False'
                     else:
                            akt = 'True'

              elif 'multilisting' in task.url:
                     if grab.doc.select(u'//meta[contains(@itemprop,"description")]').exists() == True:
                            akt = 'True'
                     else:
                            akt = 'False'

              #elif 'realt66.ru/' in task.url:
                     #if grab.doc.code ==200:
                            #akt = 'True'
                     #else:
                            #akt = 'False'

              #elif 'base.zem.ru/' in task.url:
                     #if grab.doc.select(u'//span[contains(text(),"Телефон:")]/following-sibling::span').exists() == True:
                            #akt = 'True'
                     #else:
                            #akt = 'False'

              elif 'VSE42.RU' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'brsn.ru/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'mob_sellcom' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'mob_givecom' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'advecs.com/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              #elif 'business-asset.ru/' in task.url:
                     #if grab.doc.code ==200:
                            #akt = 'True'
                     #else:
                            #akt = 'False'

              #elif 'delomart.ru/' in task.url:
                     #if grab.doc.code ==200:
                            #akt = 'True'
                     #else:
                            #akt = 'False'

              #elif 'alterainvest.ru/' in task.url:
                     #if grab.doc.code ==200:
                            #akt = 'True'
                     #else:
                            #akt = 'False'

              #elif 'biztorg.ru/' in task.url:
                     #if grab.doc.code ==200:
                            #akt = 'True'
                     #else:
                            #akt = 'False'

              elif 'БЦИнформ' in task.ist:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              elif 'bazametrov.ru/' in task.url:
                     if grab.doc.code ==200:
                            akt = 'True'
                     else:
                            akt = 'False'

              #elif 'roszem.ru/' in task.url:
                     #if grab.doc.select(u'//dt[contains(text(),"Телефон")]/following-sibling::dd').exists() == True:
                            #akt = 'True'
                     #else:
                            #akt = 'False'

              #elif 'ned72.ru/' in task.url:
                     #if grab.doc.code ==200:
                            #akt = 'True'
                     #else:
                            #akt = 'False'
              else:
                     akt =''

              self.ws.write(self.row, 3, akt)
              self.ws.write_string(self.row, 2, task.url)
              self.ws.write(self.row, 1, task.ist)
              self.ws.write(self.row, 0, task.cod)
              print('*'*50)
              print akt
              print 'Ready - '+str(self.row)+'/'+str(self.sheet.nrows)
              print 'Tasks - %s' % self.task_queue.size()
              print name
              print('*'*50)
              self.row+= 1

              #if self.row > 20:
                     #self.stop()


bot = Actuall(thread_number=80, network_try_limit=20)
bot.setup_queue('mongo', database='actualcom1',host='127.0.0.1')
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=100)
try:
       bot.run()
except KeyboardInterrupt:
       pass
print('Wait 2 sec...')
time.sleep(1)
print('Save it...')
p = os.system('echo %s|sudo -S %s' % ('1122', 'mount -a'))
print p
time.sleep(2)
bot.workbook.close()
print('Done!')


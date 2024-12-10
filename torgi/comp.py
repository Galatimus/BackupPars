#!/usr/bin/env python
# -*- coding: utf-8 -*-



from grab.spider import Spider,Task
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError 
import logging
import time
import re
from datetime import datetime
import xlsxwriter
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logger = logging.getLogger()
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


  

workbook = xlsxwriter.Workbook(u'0001-0057_00_C_001-0220_TORGI_Продажа.xlsx')

     



class Torgi_Zem(Spider):
     
     
     
     def prepare(self):
	  #self.count = 1 
	  #self.f = page
	  #self.link =l[i]
	  self.ws = workbook.add_worksheet(u'Torgi_Коммерческая_Продажа')
	  self.ws.write(0, 0, u"Организатор торгов")
	  self.ws.write(0, 1, u"Статус торгов")
	  self.ws.write(0, 2, u"Тип имущества")
	  self.ws.write(0, 3, u"Вид собственности")
	  self.ws.write(0, 4, u"Вид договора")
	  self.ws.write(0, 5, u"Целевое назначение")
	  self.ws.write(0, 6, u"Описание и технические характеристики имущества")
	  self.ws.write(0, 7, u"Страна размещения")
	  self.ws.write(0, 8, u"Субъект РФ")
	  self.ws.write(0, 9, u"Местоположение имущества")
	  self.ws.write(0, 10, u"Детальное местоположение имущества")
	  self.ws.write(0, 11, u"Площадь земельного участка")
	  self.ws.write(0, 12, u"Срок заключения договора купли-продажи")
	  self.ws.write(0, 13, u"Валюта лота")
	  self.ws.write(0, 14, u"Предмет торга")
	  self.ws.write(0, 15, u"Ежемесячная начальная цена 1 кв.м")
	  self.ws.write(0, 16, u"Ежемесячный платеж за объект")
	  self.ws.write(0, 17, u"Ежегодный платеж за объект")
	  self.ws.write(0, 18, u"Почасовой платеж за объект")
	  self.ws.write(0, 19, u"Платеж за право заключить договор")
	  self.ws.write(0, 20, u"Общая начальная (минимальная) цена за договор")
	  self.ws.write(0, 21, u"Размер задатка")
	  self.ws.write(0, 22, u"Размер обеспечения")
	  self.ws.write(0, 23, u"Субаренда")
	  self.ws.write(0, 24, u"Обременение")
	  self.ws.write(0, 25, u"Описание обременения")
	  self.ws.write(0, 26, u"Адрес")
	  self.ws.write(0, 27, u"Телефон")
	  self.ws.write(0, 28, u"Факс")
	  self.ws.write(0, 29, u"E-Mail")
	  self.ws.write(0, 30, u"Контактное лицо")
	  self.ws.write(0, 31, u"Сайт размещения документации о торгах")
	  self.ws.write(0, 32, u"Комиссия")
	  self.ws.write(0, 33, u"Только для субъектов малого и среднего предпринимательства")
	  self.ws.write(0, 34, u"Срок, место и порядок предоставления документации о торгах")
	  self.ws.write(0, 35, u"Размер платы за документацию (руб.)")
	  self.ws.write(0, 36, u"Дата окончания подачи заявок")
	  self.ws.write(0, 37, u"Срок отказа от проведения торгов")
	  self.ws.write(0, 38, u"Дата и время проведения аукциона")
	  self.ws.write(0, 39, u"Место проведения аукциона")
	  self.ws.write(0, 40, u"Место и срок подведения итогов")
	  self.ws.write(0, 41, u"Подписано электронной подписью")
	  self.ws.write(0, 42, u"Дата и время вскрытия конвертов")
	  self.ws.write(0, 43, u"Дата публикации извещения")
	  self.ws.write(0, 44, u"Дата окончания приема заявок")
	  self.ws.write(0, 45, u"Дата подведения итогов")
	  self.ws.write(0, 46, u"Место вскрытия конвертов")
	  self.ws.write(0, 47, u"Дата рассмотрения заявок")
	  self.ws.write(0, 48, u"Дата заключения договора")
	  self.ws.write(0, 49, u"Дата парсинга")
	  self.ws.write(0, 50, u"Ссылка на объект")
	  self.ws.write(0, 51, u"Описание")
	  self.ws.write(0, 52, u"Дата отмены")
	  self.ws.write(0, 53, u"Победитель торгов")
	  self.ws.write(0, 54, u"Начальная цена продажи имущества")
	  self.ws.write(0, 55, u"Цена сделки")
	  self.ws.write(0, 56, u"Наименование и характеристика имущества")
	  self.ws.write(0, 57, u"Минимальная цена")
		  
	  self.result= 1
	  
       
       
       
	 

     def task_generator(self):
	  l= open('links/Torgi_Com_prod.txt').read().splitlines()
	  self.dc = len(l)
	  print self.dc
	  for line in l:
	       yield Task ('item',url=line,refresh_cache=True,network_try_count=100)
        
            
	       
	         
	 
        
        
        
     def task_item(self, grab, task):
	  try:
	       org = grab.doc.select(u'//label[contains(text(),"Организатор торгов")]/following::span[1]').text()
	  except IndexError:
	       org=''
	  try:
	       try:
	            stat = grab.doc.select(u'//label[contains(text(),"Статус:")]/following::span[1]').text()
	       except IndexError:
		    stat = grab.doc.select(u'//label[contains(text(),"Статус торгов:")]/following::span[1]').text()
	  except IndexError:
	       stat = ''
	  try:
	       tip_im = grab.doc.select(u'//label[contains(text(),"Тип имущества:")]/following::span[1]').text()
	  except IndexError:
	       tip_im = ''
	  try:
	       vid_sob = grab.doc.select(u'//label[contains(text(),"Вид собственности:")]/following::span[1]').text()
	  except IndexError:
	       vid_sob = ''
	  try:
	       vid_dog = grab.doc.select(u'//label[contains(text(),"Вид договора:")]/following::span[1]').text()
	  except IndexError:
	       vid_dog = ''
	  try:
	       naz = grab.doc.select(u'//label[contains(text(),"Целевое назначение:")]/following::span[1]/p').text()
	  except IndexError:
	       naz = ''
	  try:
	       opis_im = grab.doc.select(u'//label[contains(text(),"Описание и технические характеристики имущества:")]/following::span[1]/p').text()
	  except IndexError:
	       opis_im = ''
	  try:
	       strana = grab.doc.select(u'//label[contains(text(),"Страна размещения:")]/following::span[1]').text()
	  except IndexError:
	       strana = ''
	  try:
	       sub = grab.doc.select(u'//label[contains(text(),"Место нахождения ")]/following::span[1]').text().split(', ')[0]
	  except IndexError:
	       sub = ''
	  try:
	       mesto = grab.doc.select(u'//label[contains(text(),"Место нахождения ")]/following::span[1]').text()
	  except IndexError:
	       mesto = ''
	  try:
	       try:
	            mesto_det = grab.doc.select(u'//label[contains(text(),"Детальное местоположение:")]/following::span[1]').text()
	       except IndexError:
	            mesto_det = grab.doc.select(u'//label[contains(text(),"Почтовый адрес:")]/following::span[1]').text()
	  except IndexError:
	       mesto_det = ''
	  try:
	       plosh = grab.doc.select(u'//label[contains(text(),"Площадь")]/following::span[@id="areaMeters"]').text()+u' м2'
	  except IndexError:
	       plosh =''
	  try:
	       srok_dog = grab.doc.select(u'//label[contains(text(),"Срок заключения договора купли-продажи:")]/following::td[1]').text()
	  except IndexError:
	       srok_dog =''
	  try:
	       val = grab.doc.select(u'//label[contains(text(),"Валюта лота:")]/following::span[1]').text()
	  except IndexError:
	       val =''
	  try:
	       pred = grab.doc.select(u'//label[contains(text(),"Предмет торга:")]/following::span[1]').text()
	  except IndexError:
	       pred =''
	  try:
	       plat_mes = grab.doc.select(u'//label[contains(text(),"Ежемесячный платеж за объект:")]/following::span[1]').text()
	  except IndexError:
	       plat_mes =''
	  try:
	       plat_god = grab.doc.select(u'//label[contains(text(),"Ежегодный платеж за объект:")]/following::span[1]').text()
	  except IndexError:
	       plat_god =''
	  try:
	       plat_chas = grab.doc.select(u'//label[contains(text(),"Почасовой платеж за объект:")]/following::span[1]').text() 
	  except IndexError:
	       plat_chas = ''
	  try:
	       plat_pravo = grab.doc.select(u'//label[contains(text(),"Платеж за право заключить договор:")]/following::span[1]').text()
	  except IndexError:
	       plat_pravo = ''
	  try:
	       ob_cena = grab.doc.select(u'//label[contains(text(),"Общая начальная (минимальная) цена за договор:")]/following::span[1]').text()
	  except IndexError:
	       ob_cena = ''
	  try:
	       raz_zad = grab.doc.select(u'//label[contains(text(),"Размер задатка:")]/following::span[1]').text()
	  except IndexError:
	       raz_zad = ''
	  try:
	       raz_ob = grab.doc.select(u'//label[contains(text(),"Размер обеспечения:")]/following::span[1]').text()
	  except IndexError:
	       raz_ob = ''
	  try:
	       sub_arenda =  grab.doc.select(u'//label[contains(text(),"Субаренда:")]/following::span[1]').text()
	  except IndexError:
	       sub_arenda = ''
	  try:
	       obrem =  grab.doc.select(u'//label[contains(text(),"Обременение:")]/following::span[1]').text()
	  except IndexError:
	       obrem = ''
	  try:
	       opis_obrem =  grab.doc.select(u'//label[contains(text(),"Описание обременения:")]/following::span[1]').text()
	  except IndexError:
	       opis_obrem = ''
	  try:
	       cena_m =  grab.doc.select(u'//label[contains(text(),"Ежемесячная начальная цена 1 кв.м:")]/following::span[1]').text()
	  except IndexError:
	       cena_m = ''
	  try:
	       data_iz =  grab.doc.select(u'//label[contains(text(),"Дата и время публикации извещения")]/following::span[1]').text()
	  except IndexError:
	       data_iz = ''
	  try:
	       data_ok =  grab.doc.select(u'//span[contains(text(),"Дата и время окончания приема заявок:")]/following::span[1]').text()
	  except IndexError:
	       data_ok = ''
          try:
               data_otm =  grab.doc.select(u'//label[contains(text(),"Дата отмены")]/following::span[1]').text()
          except IndexError:
               data_otm = ''
          try:
               pobeda =  grab.doc.select(u'//label[contains(text(),"Покупатель:")]/following::span[1]').text()
          except IndexError:
               pobeda = ''
	  try:
	       nach_cena =  grab.doc.select(u'//label[contains(text(),"Начальная цена продажи имущества:")]/following::span[1]').text()
	  except IndexError:
	       nach_cena = ''
          try:
               cena_sdel =  grab.doc.select(u'//label[contains(text(),"Цена сделки:")]/following::span[1]').text()
          except IndexError:
               cena_sdel = ''
	       
	  try:
	       min_cena =  grab.doc.select(u'//label[contains(text(),"Минимальная цена")]/following::span[1]').text()
	  except IndexError:
	       min_cena = ''	  
          try:
	       try:
                    xarakt =  grab.doc.select(u'//label[contains(text(),"Наименование и характеристики имущества:")]/following::span[1]').text()
	       except IndexError:
		    xarakt =  grab.doc.select(u'//label[contains(text(),"Наименование и характеристика имущества:")]/following::span[1]').text()
          except IndexError:
               xarakt = ''
	  ##########################################################################################################################################################     
       
         
          
	  projects = {'url': task.url,
	                 'sub': sub,
	                 'org': org,
	                 'stat': stat,
	                 'tip_im': tip_im,
	                 'vid_sob': vid_sob,
	                 'vid_dog': vid_dog,
	                 'naz': naz,
	                 'opis_im': opis_im,
	                 'strana': strana,
	                 'mesto': mesto,
	                 'mesto_det': mesto_det,
	                 'plosh': plosh,
	                 'srok_dog':srok_dog,
	                 'val': val,
	                 'pred': pred,
	                 'plat_mes': plat_mes,
	                 'plat_god': plat_god,
	                 'plat_chas': plat_chas,
	                 'plat_pravo': plat_pravo,
	                 'ob_cena':ob_cena,
	                 'raz_zad':raz_zad,
	                 'raz_ob':raz_ob,
	                 'sub_arenda': sub_arenda,
	                 'obrem': obrem,
	                 'opis_obrem': opis_obrem,
	                 'cena_m': cena_m,
	                 'data_iz': data_iz,
	                 'data_ok': data_ok,
	                 'data_otmeni': data_otm,
	                 'winer': pobeda,
	                 'cenam': min_cena,
	                 'cena1': nach_cena,
	                 'cena_zdelki': cena_sdel,
	                 'name_im': xarakt}
	                 ####################
	                
   
   
          gr = grab.clone()
          gr.setup(url='https://torgi.gov.ru/?wicket:interface=:0:notificationEditPanel:tabs:tabs-container-parent:tabs-container:tabs:0:link::IBehaviorListener:0:2')
          yield Task('next',grab=gr,project=projects,refresh_cache=True,network_try_count=100)
	  del gr
	  
	  
     def task_next(self, grab, task):
	  try:
	       adres =  grab.doc.select(u'//label[contains(text(),"Адрес:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       adres = ''
	  try:
	       phone =  grab.doc.select(u'//label[contains(text(),"Телефон:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       phone = ''
	  try:
	       fax =  grab.doc.select(u'//label[contains(text(),"Факс:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       fax = ''
	  try:
	       mail = grab.doc.select(u'//label[contains(text(),"E-Mail:")]/following::span[1]/a').text()
	  except (IndexError,AttributeError):
	       mail = ''
	  try:
	       lico = grab.doc.select(u'//label[contains(text(),"Контактное лицо:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       lico = ''

	  try:
	       side = grab.doc.select(u'//label[contains(text(),"Сайт размещения документации о торгах:")]/following::span[1]/a').text()
	  except (IndexError,AttributeError):
	       side = ''
	  try:
	       komis = grab.doc.select(u'//label[contains(text(),"Комиссия:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       komis = ''
	  try:
	       tolko = grab.doc.select(u'//label[contains(text(),"Только для субъектов малого и среднего предпринимательства:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       tolko = ''
	  try:
	       srok = grab.doc.select(u'//label[contains(text(),"Срок, место и порядок предоставления документации о торгах:")]/following::span[1]/p').text()
	  except (IndexError,AttributeError):
	       srok = ''
	  try:
	       raz_opl = grab.doc.select(u'//label[contains(text(),"Размер платы за документацию (руб.):")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       raz_opl = ''

	  try:
	       data_zaya =  grab.doc.select(u'//label[contains(text(),"Дата окончания подачи заявок:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       data_zaya = ''
	  try:
	       srok_otk = grab.doc.select(u'//label[contains(text(),"Срок отказа от проведения торгов:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       srok_otk = ''
	  try:
	       data_auk = grab.doc.select(u'//label[contains(text(),"Дата и время проведения аукциона:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       data_auk = ''
	  try:
	       mesto_auk =  grab.doc.select(u'//label[contains(text(),"Место проведения аукциона:")]/following::span[1]/p').text()
	  except (IndexError,AttributeError):
	       mesto_auk = ''
	  try:
	       mesto_itog = grab.doc.select(u'//label[contains(text(),"Место и срок подведения итогов:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       mesto_itog = ''

	  try:
	       el_pod = grab.doc.select(u'//label[contains(text(),"Подписано электронной подписью:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       el_pod = ''
	  try:
	       data_konv = grab.doc.select(u'//label[contains(text(),"Дата и время вскрытия конвертов:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       data_konv = ''
	  try:
	       data_pod_itog = grab.doc.select(u'//label[contains(text(),"Дата подведения итогов:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       data_pod_itog = ''
	  try:
	       mesto_konv = grab.doc.select(u'//label[contains(text(),"Место вскрытия конвертов:")]/following::span[1]/p').text()
	  except (IndexError,AttributeError):
	       mesto_konv = ''
	  try:
	       data_ras_zay =  grab.doc.select(u'//label[contains(text(),"Дата рассмотрения заявок:")]/following::span[1]').text()
	  except (IndexError,AttributeError):
	       data_ras_zay = ''	  
	  
	  project2 = {'adress': adres,
	              'phone': phone,
	              'fax': fax,
	              'mail': mail,
	              'lico': lico,
	              'side': side,
	              'komis': komis,
	              'tolko': tolko,
	              'srok': srok,
	              'raz_opl': raz_opl,
	              'data_zaya': data_zaya,
	              'srok_otk': srok_otk,
	              'data_auk':data_auk,
	              'mesto_auk': mesto_auk,
	              'mesto_itog': mesto_itog,
	              'el_pod': el_pod,
	              'data_konv': data_konv,
	              'data_pod_itog': data_pod_itog,
	              'mesto_konv': mesto_konv,
	              'data_ras_zay':data_ras_zay}
	  
	  
	  
	  
	  
	  
	  
	  yield Task('write',project=task.project,proj=project2,grab=grab)
	

	
	
	
	
     def task_write(self,grab,task):
	
	  print('*'*100)
	  print  task.project['org']
	  print  task.project['stat']
	  print  task.project['tip_im']
	  print  task.project['vid_sob']
	  print  task.project['vid_dog']
	  print  task.project['naz']
	  print  task.project['sub']
	  print  task.project['opis_im']
	  print  task.project['strana']
	  print  task.project['mesto']
	  print  task.project['mesto_det']
	  print  task.project['plosh']
	  print  task.project['srok_dog']
	  print  task.project['val']
	  print  task.project['pred']
	  print  task.project['plat_mes']
	  print  task.project['plat_god']
	  print  task.project['plat_chas']
	  print task.project['plat_pravo']
	  print  task.project['ob_cena']
	  print  task.project['raz_zad']
	  print  task.project['raz_ob']
	  print  task.project['sub_arenda']
	  print  task.project['obrem']
	  print  task.project['opis_obrem']
	  print  task.project['cena_m']
	  print  task.project['data_iz']
	  print  task.project['data_ok']
	  print  task.project['data_otmeni']
	  print  task.project['winer']
	  print  task.project['cena1']
	  print  task.project['cena_zdelki']
	  print  task.project['cenam']
	  
	  #############################
	  print('*'*50)
	  print  task.proj['adress']
	  print  task.proj['phone']
	  print  task.proj['fax']
	  print  task.proj['mail']
	  print  task.proj['lico']
	  print  task.proj['side']
	  print  task.proj['komis']
	  print  task.proj['tolko']
	  print  task.proj['srok']
	  print  task.proj['raz_opl']
	  print  task.proj['data_zaya']
	  print  task.proj['srok_otk']
	  print  task.proj['data_auk']
	  print  task.proj['mesto_auk']
	  print  task.proj['mesto_itog']
	  print  task.proj['el_pod']
	  print  task.proj['data_konv']
	  print  task.proj['data_pod_itog']
	  print  task.proj['mesto_konv']
	  print  task.proj['data_ras_zay']
	  print  task.project['url']
	  print  task.project['name_im']
	  
	  
	
	  
	  
	  self.ws.write(self.result, 0, task.project['org'])
	  self.ws.write(self.result, 1, task.project['stat'])
	  self.ws.write(self.result, 2, task.project['tip_im'])
	  self.ws.write(self.result, 3, task.project['vid_sob'])
	  self.ws.write(self.result, 4, task.project['vid_dog'])
	  self.ws.write(self.result, 5, task.project['naz'])
	  self.ws.write(self.result, 6, task.project['opis_im'])
	  self.ws.write(self.result, 7, task.project['strana'])
	  self.ws.write(self.result, 8, task.project['sub'])
	  self.ws.write(self.result, 9, task.project['mesto'])
	  self.ws.write(self.result, 10, task.project['mesto_det'])
	  self.ws.write(self.result, 11, task.project['plosh'])
	  self.ws.write(self.result, 12, task.project['srok_dog'])
	  self.ws.write(self.result, 13, task.project['val'])
	  self.ws.write(self.result, 14, task.project['pred'])
	  self.ws.write(self.result, 15, task.project['cena_m'])
	  
	  self.ws.write(self.result, 16, task.project['plat_mes'])
	  self.ws.write(self.result, 17, task.project['plat_god'])
	  self.ws.write(self.result, 18, task.project['plat_chas'])
	  self.ws.write(self.result, 19, task.project['plat_pravo'])
	  self.ws.write(self.result, 20, task.project['ob_cena'])
	  self.ws.write(self.result, 21, task.project['raz_zad'])
	  self.ws.write(self.result, 22, task.project['raz_ob'])
	  self.ws.write(self.result, 23, task.project['sub_arenda'])
	  self.ws.write(self.result, 24, task.project['obrem'])
	  self.ws.write(self.result, 25, task.project['opis_obrem'])
	  self.ws.write(self.result, 26, task.proj['adress'])
	  self.ws.write(self.result, 27, task.proj['phone'])
	  self.ws.write(self.result, 28, task.proj['fax'])
	  self.ws.write(self.result, 29, task.proj['mail'])
	  self.ws.write(self.result, 30, task.proj['lico'])
	  self.ws.write_string(self.result, 31, task.proj['side'])
	  self.ws.write(self.result, 32, task.proj['komis'])
	  self.ws.write(self.result, 33, task.proj['tolko'])
	  self.ws.write(self.result, 34, task.proj['srok'])
	  self.ws.write(self.result, 35, task.proj['raz_opl'])
	  self.ws.write(self.result, 36, task.proj['data_zaya'])
	  self.ws.write(self.result, 37, task.proj['srok_otk'])
	  self.ws.write(self.result, 38, task.proj['data_auk'])
	  self.ws.write(self.result, 39, task.proj['mesto_auk'])
	  self.ws.write(self.result, 40, task.proj['mesto_itog'])
	  self.ws.write(self.result, 41, task.proj['el_pod'])
	  
	  self.ws.write(self.result, 42, task.proj['data_konv'])
	  self.ws.write(self.result, 43, task.project['data_iz'])
	  self.ws.write(self.result, 44, task.project['data_ok'])
	  self.ws.write(self.result, 45, task.proj['data_pod_itog'])
	  self.ws.write(self.result, 46, task.proj['mesto_konv'])
	  self.ws.write(self.result, 47, task.proj['data_ras_zay'])
	  self.ws.write(self.result, 49, datetime.today().strftime('%d.%m.%Y'))
	  self.ws.write_string(self.result, 50, task.project['url'])
	  self.ws.write(self.result, 51, 'Организатор торгов - '+ task.project['org']+','
              +'Вид собственности - '+ task.project['vid_sob']+','
              +'Срок заключения договора - '+ task.project['srok_dog']+','
              +'Платеж за право заключить договор - '+ task.project['plat_pravo']+','
              +'Обременение - '+ task.project['obrem']+','
              +'Описание обременения - '+ task.project['opis_obrem']+','
              +'Комиссия - '+ task.proj['komis']+','
              +'Только для субъектов малого и среднего предпринимательства - '+ task.proj['tolko']+','
              +'Срок отказа от проведения торгов - '+ task.proj['srok_otk']+','
              +'Дата и время проведения аукциона - '+ task.proj['data_auk']+','
              +'Ежемесячная начальная цена 1 кв.м -'+ task.project['cena_m']+','
              +'Дата окончания подачи заявок - '+ task.project['data_ok'])
	  self.ws.write(self.result, 52, task.project['data_otmeni'])
	  self.ws.write(self.result, 53, task.project['winer'])
	  self.ws.write(self.result, 54, task.project['cena1'])
	  self.ws.write(self.result, 55, task.project['cena_zdelki'])
	  self.ws.write(self.result, 56, task.project['name_im'])
	  self.ws.write(self.result, 57, task.project['cenam'])
	 
	 
	  
   
	  print('*'*100)
	  
	  print 'Ready - '+str(self.result)+'/'+str(self.dc)
	  logger.debug('Tasks - %s' % self.task_queue.size()) 
	  #print '***',i+1,'/',dc,'***'
	  print('*'*100)
	  	  
	  self.result+= 1
	  
	  #if self.result > 10:
	       #self.stop()
	

bot = Torgi_Zem(thread_number=5, network_try_limit=1000)
bot.load_proxylist('../tipa.txt','text_file')
bot.create_grab_instance(timeout=50, connect_timeout=50)
bot.run()
print('Wait 2 sec...')
time.sleep(2)
print('Save it...')    
command = 'mount -a'
p = os.system('echo %s|sudo -S %s' % ('1122', command))
print p
time.sleep(1)
workbook.close()
print('Done!')

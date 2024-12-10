#!/usr/bin/python
# -*- coding: utf-8 -*-



import math
import os
import re
import time
from datetime import datetime,timedelta
from sub import conv
import sys
import xlsxwriter
from selenium import webdriver
from selenium.common.exceptions import TimeoutException,WebDriverException,NoSuchElementException


reload(sys)
sys.setdefaultencoding('utf-8')



def get_chrome_drive(driver_path=None):
    
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument('--hide-scrollbars')
    options.add_argument('--no-sandbox') 
    driver = webdriver.Chrome(executable_path='D:\\VMF\\OlegPars\\webshot\\chromedriver\\chromedriver.exe',chrome_options=options,service_args=['--verbose']) 
    return driver

def get_firefox_drive(driver_path=None):

    options = webdriver.FirefoxOptions()
    options.add_argument('-headless')
    profile =  webdriver.FirefoxProfile('/home/oleg/.mozilla/firefox/46stx7t7.default/') #Gui1
    #profile = webdriver.FirefoxProfile()#Gui2
    profile.set_preference('permissions.default.stylesheet', 2)
    profile.set_preference('permissions.default.image', 2)
    profile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', False)
    profile.set_preference("javascript.enabled", False)
    profile.native_events_enabled = False
    driver = webdriver.Firefox(executable_path=r'/usr/local/bin/geckodriver',firefox_profile=profile,firefox_options=options,service_log_path=None)
    return driver

def take_content(driver, url):
    # get the page
    driver.get(url)
    time.sleep(3)
    try:
        zag = driver.title
    except (NoSuchElementException,WebDriverException):
        zag = ''
    try:
        uliza = driver.find_element_by_xpath('//div[@class="street"]/a').text
    except (NoSuchElementException,IndexError,WebDriverException):
        uliza = ''
    try:
        ray = driver.find_element_by_xpath('//div[@class="district"]/a').text
    except (NoSuchElementException,IndexError,WebDriverException):
        ray = ''
    try:
        if 'moscow' in url:
            punkt = u'Москва'
        else:        
            punkt = driver.find_element_by_xpath('//ul[@class="breadcrumbs"]/li[2]/span/a/span[1]').text
    except (NoSuchElementException,IndexError,WebDriverException):
        punkt =''
    try:
        cena = driver.find_element_by_xpath('//td[@itemprop="price"][1]').text
    except (NoSuchElementException,IndexError,WebDriverException):
        cena = ''
    try:
        oren = driver.find_element_by_xpath('//span[@class="left"]/a').text
    except (NoSuchElementException,IndexError,WebDriverException):
        oren = ''
    try:
        seg = driver.find_element_by_xpath('//tr[@class="object-id"]/td[2]').text
    except (NoSuchElementException,IndexError,WebDriverException):
        seg =''
    try:
        klass = driver.find_element_by_xpath('//span[@class="white-text badge"]').text
    except (NoSuchElementException,IndexError,WebDriverException):
        klass = ''
    try:
        try:
            plosh = driver.find_element_by_xpath('//ul[@class="breadcrumbs"]/li[4]/descendant::span[2]').text
        except (NoSuchElementException,IndexError,WebDriverException):
            plosh = re.sub('[^\d\.]','',driver.find_element_by_xpath(u'//meta[@name="description"]').get_attribute('content').split(': ')[1].split(u' за ')[0].split(u' кв.м')[0])+' м2'
    except (NoSuchElementException,IndexError,WebDriverException):
        plosh = ''
    try:
        ets = driver.find_element_by_xpath('//tr[@class="object-id"]/td[3]').text
    except (NoSuchElementException,IndexError,WebDriverException):
        ets = ''
    try:
        metro = driver.find_element_by_xpath('//span[@class="metro-line"]/following-sibling::text()').text.split('(')[0]
    except (NoSuchElementException,IndexError,WebDriverException):
        metro = ''
    try:
        opis = driver.find_element_by_xpath('//div[@class="extended-body"]').text
    except (NoSuchElementException,IndexError,WebDriverException):
        opis = ''
    try:
        lico = driver.find_element_by_xpath('//div[@class="name"]').text
    except (NoSuchElementException,IndexError,WebDriverException):
        lico = ''
    try:
        phone = driver.find_element_by_xpath('//div[@class="phone"]').get_attribute('data-phone')
    except (NoSuchElementException,IndexError,WebDriverException):
        phone =''
    try:
        data = driver.find_element_by_xpath('//div[@class="row"]/div[contains(@class,"lastModify")]').text
    except (NoSuchElementException,IndexError,WebDriverException):
        data = ''
    try:
        #oper = driver.find_element_by_xpath('//ul[@class="breadcrumbs"]/li[3]/descendant::span[2]/text()')[0].split(' ')[0]
        oper = driver.find_element_by_xpath(u'//meta[@name="description"]').get_attribute('content').split(': ')[1].split(' ')[0]
    except (NoSuchElementException,IndexError,WebDriverException):
        oper = ''
        
    sub = reduce(lambda punkt, r: punkt.replace(r[0], r[1]), conv, punkt)
    
    data = re.sub(u"[^а-яА-Я0-9.,\-\s]", "", data)
    data = re.sub(u"[.,\-\s]{3,}", " ", data).replace(u'Данные обновлены ','').replace('-','.')[1:].split(' ')[0]    
        
    print('*'*50)
    print sub 
    print punkt 
    print ray 
    print uliza
    print oren
    print seg
    print klass
    print cena
    print plosh
    print ets
    print opis
    print phone
    print lico
    print oper
    print data
    print metro
    print zag
    print('*'*50)
    ws.write(result, 0, sub)
    ws.write(result, 1, ray)
    ws.write(result, 2, punkt)
    ws.write(result, 4, uliza)
    ws.write(result, 6, oren)
    ws.write(result, 7, seg)
    ws.write(result, 10, klass)
    ws.write(result, 11, cena)
    ws.write(result, 14, plosh)
    ws.write(result, 16, ets)
    ws.write(result, 18, opis)
    ws.write(result, 19, u'БЦИнформ')
    ws.write_string(result, 20, url)
    ws.write(result, 21, phone)
    ws.write(result, 22, lico)
    ws.write(result, 26, metro)
    ws.write(result, 28, oper)
    ws.write(result, 30, data)
    ws.write(result, 31, datetime.today().strftime('%d.%m.%Y'))
    ws.write(result, 33, zag)   


def main(url):
    driver = get_firefox_drive()
    #driver = get_chrome_drive()

    driver.set_window_size(800,800)
    try:
        take_content(driver,url)
    except (TimeoutException,WebDriverException):
        driver.quit()
    
    driver.quit()
    time.sleep(1)    
    return


if __name__ == '__main__':
    l= open('caos.txt').read().splitlines()
    workbook = xlsxwriter.Workbook(u'0001-0002_00_C_001-0217_BCINF.xlsx')    
    ws = workbook.add_worksheet()
    ws.write(0, 0, u"СУБЪЕКТ_РОССИЙСКОЙ_ФЕДЕРАЦИИ")
    ws.write(0, 1, u"МУНИЦИПАЛЬНОЕ_ОБРАЗОВАНИЕ_(РАЙОН)")
    ws.write(0, 2, u"НАСЕЛЕННЫЙ_ПУНКТ")
    ws.write(0, 3, u"ВНУТРИГОРОДСКАЯ_ТЕРРИТОРИЯ")
    ws.write(0, 4, u"УЛИЦА")
    ws.write(0, 5, u"ДОМ")
    ws.write(0, 6, u"ОРИЕНТИР")
    ws.write(0, 7, u"СЕГМЕНТ")
    ws.write(0, 8, u"ТИП_ПОСТРОЙКИ")
    ws.write(0, 9, u"НАЗНАЧЕНИЕ_ОБЪЕКТА")
    ws.write(0, 10, u"ПОТРЕБИТЕЛЬСКИЙ_КЛАСС")
    ws.write(0, 11, u"СТОИМОСТЬ")
    ws.write(0, 12, u"ИЗМЕНЕНИЕ_СТОИМОСТИ")
    ws.write(0, 13, u"ДОПОЛНИТЕЛЬНЫЕ_КОММЕРЧЕСКИЕ_УСЛОВИЯ")
    ws.write(0, 14, u"ПЛОЩАДЬ")
    ws.write(0, 15, u"ЭТАЖ")
    ws.write(0, 16, u"ЭТАЖНОСТЬ")
    ws.write(0, 17, u"ГОД_ПОСТРОЙКИ")
    ws.write(0, 18, u"ОПИСАНИЕ")
    ws.write(0, 19, u"ИСТОЧНИК_ИНФОРМАЦИИ")
    ws.write(0, 20, u"ССЫЛКА_НА_ИСТОЧНИК_ИНФОРМАЦИИ")
    ws.write(0, 21, u"ТЕЛЕФОН")
    ws.write(0, 22, u"КОНТАКТНОЕ_ЛИЦО")
    ws.write(0, 23, u"КОМПАНИЯ")
    ws.write(0, 24, u"МЕСТОПОЛОЖЕНИЕ")
    ws.write(0, 25, u"МЕСТОРАСПОЛОЖЕНИЕ")
    ws.write(0, 26, u"БЛИЖАЙШАЯ_СТАНЦИЯ_МЕТРО")
    ws.write(0, 27, u"РАССТОЯНИЕ_ДО_БЛИЖАЙШЕЙ_СТАНЦИИ_МЕТРО")
    ws.write(0, 28, u"ОПЕРАЦИЯ")
    ws.write(0, 29, u"ДАТА_РАЗМЕЩЕНИЯ")
    ws.write(0, 30, u"ДАТА_ОБНОВЛЕНИЯ")
    ws.write(0, 31, u"ДАТА_ПАРСИНГА")
    ws.write(0, 32, u"КАДАСТРОВЫЙ_НОМЕР")
    ws.write(0, 33, u"ЗАГОЛОВОК")
    ws.write(0, 34, u"ШИРОТА_ИСХ")
    ws.write(0, 35, u"ДОЛГОТА_ИСХ")
    result= 1
    try:
        for p in range(len(l)):
            print '******',p,'/',len(l),'******'
            main(l[p])
            result+=1  
    except KeyboardInterrupt:
        pass
    print('Save it...')
    time.sleep(1)
    workbook.close()
    time.sleep(2)
    print('Done')    
    
    
    

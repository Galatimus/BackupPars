#!/usr/bin/python
# -*- coding: utf-8 -*-


from grab import Grab
import time
import logging
import os
from grab.error import GrabTimeoutError, GrabNetworkError,DataNotFound,GrabConnectionError
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)

i = 0
l= open('city_cian.txt').read().splitlines()
page ='snyat-pomeshenie-v-biznes-centre/'


headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
         #'Accept-Encoding': 'gzip,deflate',
        'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
        #'Cookie': '_CIAN_GK=728644fe-916e-416e-af6d-41d9bf52f7cf',
        #'Host': 'qp.ru',
        #'Referer': task.url,
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0'}



while True:
    print '**********',i+1,'/',len(l),'********'
    try:
        time.sleep(2)
        g = Grab(timeout=20, connect_timeout=50)
        g.proxylist.load_file(path='../ivan.txt',proxy_type='http')
        #g.go(l[i]+page,headers=headers)
        g.request(headers=headers,url=l[i]+page)
        print g.doc.code
       
    except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
        print g.config['proxy'],'Change proxy'
        g.change_proxy()
        del g
        continue
    
    print g.config['url']
    print g.config['proxy']
    
    if g.doc.code <> 200 :    
        time.sleep(2)
        continue     
        
    time.sleep(1)
    lin = []
    #del g

    while True:
        try:
            time.sleep(2)   
            g2 = g.clone(headers=headers,proxy_auto_change=True)
            #g = Grab(timeout=20, connect_timeout=50)
            #g.proxylist.load_file(path='../tipa.txt',proxy_type='http')
            g2.request(headers=headers,url=l[i]+page)
            for link in g2.doc.select('//h3/a'):
                url = g2.make_url_absolute(link.attr('href'))
                print url
                lin.append(url)
            print '***',len(lin),'**********',i+1,'/',len(l),'********'
            time.sleep(1)   
            print "Next Page is ..."  
            nextpage =  g2.make_url_absolute(g.doc.select(u'//nav[@class="cf-pagination"]/span/following-sibling::a[1]').attr('href'))
            time.sleep(2)
            g2.request(headers=headers,url=nextpage)
            #url_next = l[i]+page+'page-'+nextpage+'/'
            #print url_next
            #print '*********************************'
            time.sleep(2)             
        except(GrabTimeoutError,GrabNetworkError,DataNotFound,GrabConnectionError):
            print g.config['proxy'],'Change proxy'
            print g.config['url']
            g.change_proxy()
            time.sleep(2)
            del g2
            continue

            #links = open('bc_com.txt', 'a')
            #for item in lin:
                #links.write("%s\n" % item)
            #links.close()            
            #time.sleep(1)            
            #print'NEXT'            
            #break
        
    #sess.reset()
    i=i+1 
    
    





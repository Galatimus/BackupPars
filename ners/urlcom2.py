#!/usr/bin/python
# -*- coding: utf-8 -*-


import dryscrape
import time
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


i = 0
l= open('Links/com_all.txt').read().splitlines()
dryscrape.start_xvfb()
sess = dryscrape.Session()
while True:
    
    sess.set_header('user-agent', 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0')
    sess.visit(l[i])
    
    time.sleep(1)
    lin = []
    while True:
        try:
            time.sleep(1)   
            for link in sess.xpath('//div[@class="media-body"]/div/following-sibling::a'):
                url = link['href']
                print url
                lin.append(url)
            print '***',len(lin),'**********',i+1,'/',len(l),'********'
            time.sleep(1)   
            print "Next Page is ..."  
            nextpage = sess.at_xpath(u'//div[@class="pagination"]/span/strong/following-sibling::a[1]')['href']
            print nextpage
            print '*********************************'
            #nextpage.click()
            sess.visit(nextpage)
            time.sleep(2)             
        except :
            links = open('bc_com2.txt', 'a')
            for item in lin:
                links.write("%s\n" % item)
            links.close()            
            time.sleep(1)            
            print'NEXT'            
            break
        
    sess.reset()
    i=i+1 
    
    





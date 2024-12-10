#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import time
import random
from lxml import html
from lxml.etree import ParserError
from lxml.etree import XMLSyntaxError
import logging
import subprocess
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(level=logging.DEBUG)
l= open('faprost_com.txt').read().splitlines()
for p in range(len(l)):
    print '******',p+1,'/',len(l),'*******'
    proxy = random.choice(list(open('../ivan.txt').read().splitlines())).split(':')[0]+':8080'
    print proxy
    address = l[p]
    command = "phantomjs --proxy %s --proxy-auth %s --proxy-type http web.js %s %s" % (proxy,'Ivan:tempuvefy','muu.png',address)
    proc = subprocess.Popen(command, shell=True,stdout=subprocess.PIPE).communicate()
    try:
        parsed_body = html.fromstring(proc[0].decode('utf-8').strip())
    except (ParserError,XMLSyntaxError):
        time.sleep(2)
        continue    
    try:
        zag = parsed_body.xpath('//title/text()')[0]
    except IndexError:
        zag = ''
    print zag
    #os.system('python webscreenshot.py -v '+l[p]+' -v'+' -P '+proxy+':8080'+' -A Ivan:tempuvefy')
    #os.system('python webscreenshot.py -v '+l[p]+' -v')
    time.sleep(5)
    
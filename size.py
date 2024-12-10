#!/usr/bin/env python
# -*- coding: utf-8 -*-



import os
import time


delsize =['/home/oleg/pars/avito/com',
            '/home/oleg/pars/avito/zem',
            '/home/oleg/pars/avito/zagg',
            '/home/oleg/pars/avito/flats',
            '/home/oleg/pars/bn/com',
            '/home/oleg/pars/bn/zem',
            '/home/oleg/pars/n1/flat',
            '/home/oleg/pars/mirkv/flat',
            '/home/oleg/pars/gde/com',
            '/home/oleg/pars/gdedom/com',
            '/home/oleg/pars/gdedom/zem',
            '/home/oleg/pars/gdedom/flats',
            '/home/oleg/pars/house/com',
            '/home/oleg/pars/house/Kv',
            '/home/oleg/pars/house/zem',
            '/home/oleg/pars/irr/Com',
            '/home/oleg/pars/irr/Zem',
            '/home/oleg/pars/irr/flats',
            '/home/oleg/pars/irr/zagg',
            '/home/oleg/pars/kvadrat/com',
            '/home/oleg/pars/kvadrat/zem',
            '/home/oleg/pars/life/com',
            '/home/oleg/pars/life/zem',
            '/home/oleg/pars/mag/com',
            '/home/oleg/pars/mag/zem',
            '/home/oleg/pars/mag/flats',
            '/home/oleg/pars/mirkv/com',
            '/home/oleg/pars/mirkv/zagg',
            '/home/oleg/pars/mirkv/zem',
            '/home/oleg/pars/mlsn/com',
            '/home/oleg/pars/mlsn/zem',
            '/home/oleg/pars/move/com',
            '/home/oleg/pars/move/zem',
            '/home/oleg/pars/move/flat',
            '/home/oleg/pars/move/zagg',
            '/home/oleg/pars/n1/com',
            '/home/oleg/pars/tvoy/com',
            '/home/oleg/pars/ners/zem',
            '/home/oleg/pars/ners/com',
            #'/home/oleg/pars/ngs/com',
            #'/home/oleg/pars/ngs/zem',
            '/home/oleg/pars/nmls/com',
            '/home/oleg/pars/nmls/zem',
            '/home/oleg/pars/nndv/com',
            '/home/oleg/pars/nndv/zem',
            '/home/oleg/pars/property/com',
            '/home/oleg/pars/property/zem',
            '/home/oleg/pars/qp/com',
            '/home/oleg/pars/qp/zem',
            '/home/oleg/pars/yand/com',
            '/home/oleg/pars/yand/zem',            
            '/home/oleg/pars/nedv/zem',
            '/home/oleg/pars/raui/com',
            '/home/oleg/pars/raui/zem',
            '/home/oleg/pars/rosnedv/com',
            '/home/oleg/pars/rosnedv/com',
            '/home/oleg/pars/rosnedv/zem',
            '/home/oleg/pars/small/aren',
            '/home/oleg/pars/rosrealt/zem',
            '/home/oleg/pars/rosrealt/com',
            '/home/oleg/pars/rosrealt/flats',
            '/home/oleg/pars/vestum/com',
            '/home/oleg/pars/vestum/zem',
            '/home/oleg/pars/vision/com',
            '/home/oleg/pars/vision/zem',
            '/home/oleg/pars/yula/com',
            '/home/oleg/pars/yula/zem',
            '/home/oleg/pars/biz/avito',
            '/home/oleg/pars/biz/Bfs',
            '/home/oleg/pars/biz/zona',
            '/home/oleg/pars/biz/yula']

i = 0

page = delsize[i]

while True:

         for folderName, subfolders, filenames in os.walk(page):
                  for filename in filenames :
                           time.sleep (0.005)
                           fileAbsPath = os.path.abspath(os.path.join(folderName, filename))
                           size = os.path.getsize(fileAbsPath)                          
                           if size < 6200:
                                    os.remove(os.path.join(fileAbsPath))
                                    print 'Удаляем: '+ str(filename) + ' Размер: = '+str(size)
                           else:
                                    print '*****',i+1,'/',len(delsize),'*****',fileAbsPath
         i=i+1
         try:
                  page = delsize[i]
         except IndexError: 
                  break                  

print 'ГОТОВО!!!'

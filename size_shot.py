#!/usr/bin/env python
# -*- coding: utf-8 -*-



import os
import time


page = '/home/oleg/pars/over/shots/shot'
   

for folderName, subfolders, filenames in os.walk(page):
         for filename in filenames :
                  #time.sleep (0.005)
                  fileAbsPath = os.path.abspath(os.path.join(folderName, filename))
                  size = os.path.getsize(fileAbsPath)
                  if size < 120000:
                           os.remove(os.path.join(fileAbsPath))
                           print 'Удаляем: '+ str(filename) + ' Размер: = '+str(size)                  
                  else:
                           print '*****',fileAbsPath,'*****'  

print 'ГОТОВО!!!'

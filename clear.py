#!/usr/bin/env python
# -*- coding: utf-8 -*-



import os
import time
from os import listdir
from delfiles import delbiz,delact,delzagkv,delcomzem
import sys




pars= ['КН и Землю','Бизнес','Актуальность','Жилье и Загород','Пустые файлы','Скрины размер']
b = 1   
link = 0
for line in pars:
        print b,'-',line
        b += 1
        
link =int(input('Что удаляем?: '))

print link

i = 0
if link == 1:
        l = delcomzem
        print 'Удаляем: КН и Землю'
elif link == 2:
        l = delbiz
        print ' Удаляем: Бизнес'
elif link == 3:
        l = delact
        print 'Удаляем: Актуальность'
elif link == 4:
        l= delzagkv
        print 'Удаляем: Жилье и Загород'
        
elif link == 5:
        os.system("/home/oleg/pars/size.py")
        exit()
elif link == 6:
        os.system("/home/oleg/pars/size_shot.py")
        exit()
        
page = l[i]

while True:
        print '*****',i+1,'/',len(l),'*****'
        for file in os.listdir(page):
                if file.endswith(".xlsx"):
                        time.sleep (0.01)
                        #print os.path.join(page, file)
                        os.remove(os.path.join(page, file))
                        print 'Удаляем: '+ str(file)                        
        time.sleep (0.01)     
        i=i+1
        try:
                page = l[i]
        except IndexError: 
                break
        
print 'ГОТОВО!!!'



#for root, dirs, files in os.walk("/home/oleg/pars/house"):
        #for file in files:
                #if file.endswith(".xlsx"):
                        ##print os.path.join(root, file)
                        #time.sleep (0.03)
                        #os.remove(os.path.join(root, file))
                        #print 'Удалено '+ str(file)
##print 'Удалено'+ str(file)


#check_list = ('com','zem')
#folder = '/home/oleg/pars/'
#need_check = []
#need_file = []
#for root, dirs, files in os.walk(folder):
        #for dirs in dirs:
                #if not dirs.endswith(check_list):
                        #print os.path.join(root, dirs)
                        #need_check.append(os.path.join(root, dirs))
                        
                        
#links = open('test.txt', 'w')
#for item in need_check:
        #links.write("%s\n" % item)
#links.close()

#time.sleep (1)
#print len(need_check)
#time.sleep (1)



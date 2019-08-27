# -*- coding: utf-8 -*-
import sys
import os
import shutil

namelist = []
nametext = 'H:\\赋能起航志愿者证书\\August\\NameList.txt'

# 从TXT文件中读取姓名列表,需要注意的是每行会作为一个列表存进去
with open(nametext, 'r') as txt:
    for line in txt:
        namelist.append(list(line.strip('\n').split(',')))

# 从列表中取出重复的姓名，并统计他们出现的次数
namedict = {}
for name in namelist:
    if namelist.count(name) > 1 :
        namedict[name[0]] = namelist.count(name)


print(namedict)

srcFolder = "H:\\赋能起航志愿者证书\\August\\JPG\\"
for file in os.listdir(srcFolder):
    srcFile = os.path.join(srcFolder, file)
    if os.path.isfile(srcFile) == True:
        filename = os.path.splitext(file)[0]
        if filename in namedict.keys():
            for i in range(2, namedict[filename]+1):
                newFile = srcFolder + filename + "_" + str(i) + ".jpg"
                shutil.copy(srcFile, newFile)
            os.rename(srcFile, srcFolder + filename + "_1.jpg")

print("Finished!")

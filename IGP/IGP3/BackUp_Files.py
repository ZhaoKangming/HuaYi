# -*- coding: utf-8 -*-

import datetime
import os
import shutil
import time

dir_list = ["原始报告", "合格报告", "原始病例", "合格病例"]
for i in range(4):
    srcdir = 'C:\\Users\\ZhaoKangming\\OneDrive - cnu.edu.cn\\文档\\华医网\\赋能起航\\报告病例审核\\' + dir_list[i]
    dstdir = 'C:\\Users\\ZhaoKangming\\OneDrive - cnu.edu.cn\\桌面\\赋能起航报告病例审核-' + str(datetime.date.today()).replace("-", "") + "\\" + dir_list[i]
    shutil.copytree(srcdir, dstdir)

    for file in os.listdir(dstdir):
        if os.path.isfile(os.path.join(dstdir, file)) == True:
            os.rename(os.path.join(dstdir, file), os.path.join(dstdir, file.split("_",1)[1]))

print("Finished!")

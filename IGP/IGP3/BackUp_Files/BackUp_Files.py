# -*- coding: utf-8 -*-

import easygui as g
import configparser
import datetime
import os
import shutil
import time

config = configparser.ConfigParser()
config.read_file(open('config.ini'))

src_folder_path = config.get("PATH", "SRC_PATH")
dst_folder_path = config.get("PATH", "DST_PATH")

dir_list = ["原始报告", "合格报告", "原始病例", "合格病例"]
for i in range(4):
    srcdir = src_folder_path + dir_list[i]
    dstdir = dst_folder_path + '赋能起航报告病例审核-' + \
        str(datetime.date.today()).replace("-", "") + "\\" + dir_list[i]
    shutil.copytree(srcdir, dstdir)

    for file in os.listdir(dstdir):
        if os.path.isfile(os.path.join(dstdir, file)) == True:
            os.rename(os.path.join(dstdir, file), os.path.join(
                dstdir, file.split("_", 1)[1]))

Finish_msg = g.msgbox(msg="所有报告病例都已经重命名好了！", title="IGP3_BackUp_Files", ok_button="Nice")

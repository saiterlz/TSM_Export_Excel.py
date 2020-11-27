# -*- coding=utf-8 -*-

import os
import configparser

# 当前文件路径
proDir = os.path.split(os.path.realpath(__file__))[0]
# 在当前文件路径下查找.ini文件
configPath = os.path.join(proDir, "config.ini")
print(configPath)
conf = configparser.ConfigParser()
# 读取.ini文件
conf.read(configPath,encoding="utf-8-sig")
# get()函数读取section里的参数值
name = conf.get('section1', 'path_excel')
print(name)
print(conf.sections())
print(conf.options('section1'))
print(conf.items('section1'))

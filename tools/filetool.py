
import os
import shutil
from tools.strtool import contains

def createFileIfnotExist(filename):
    if not checkfileexist(filename):
        createfile(filename)

#创建新文件
def createfile(filename,value=""):
    if checkfileexist(filename):
        print("文件已经存在！！！！")
        return
    print("文件不存在！！！！")
    dirpath = os.path.split(filename)[0]
    print("dirpath:"+dirpath)
    if not checkPathexist(dirpath):
        print("目录不存在")
        os.mkdir(dirpath)
    fp=open(filename, "w+")
    fp.write(value)
    fp.close()

#创建新目录
def createdir(dirname):
    if checkPathexist(dirname):
        return
    os.makedirs(dirname)

# 扫描文件,并将路径+文件名以列表形式返回
def searchForFile(path,splix):
    items=[]
    dirs = os.listdir(path)
    for file in dirs:
        filepath = path + "\\" + str(file)
        if os.path.isdir(filepath):
            items+=searchForFile(filepath,splix)
        if os.path.isfile(filepath):
            if os.path.splitext(filepath)[1]== splix:
                print(filepath)
                items.append(filepath)
    return items

#返回当前路径目录
def get_local_dir(path):
    dirs=os.listdir(path)
    return dirs


class FileDispatcher():
    def __init__(self):
        pass
    def file_dispatch(filepath):
        return filepath

from functools import singledispatch
from collections import abc
@singledispatch
def searchFromPath(obj):
    content = repr(obj)
    return content
@searchFromPath.register(str)
def _(strPath):
    return strPath

@searchFromPath.register(tuple)
@searchFromPath.register(abc.MutableSequence)
def _(seq):
    inner = '</li>\n<li>'.join(searchFromPath(item) for item in seq)
    return '<ul>\n<li>' + inner + '</li>\n</ul>'




#检索符合条件的文件并处理，将结果返回列表items
def searchforFileWithCallback(path,dper):
    items = []
    if os.path.isdir(path):
        dirs = os.listdir(path)
        for file in dirs:
            _filepath = path + "\\" + str(file)
            if os.path.isdir(_filepath):
                item=searchforFileWithCallback(path=_filepath,dper=dper)
                if item is not None:
                    items += item
            if os.path.isfile(_filepath):
                item=dper.file_dispatch(filepath=_filepath)
                if item is not None and len(item)>0:
                    items.append(item)
    elif os.path.isfile(path):
        item = dper.file_dispatch(filepath=path)
        if len(item) > 0:
            items.append(item)
    else:
        return None
    return items

#复制文件到新的地址
def copyFile(oldfile,newfile):
    #如果旧文件不存在
    if not checkfileexist(oldfile):
        print("target file didn't exist!")
        return
    #如果新文件目录不存在
    newfiledir=os.path.split(newfile)[0]
    if not os.path.isdir(newfiledir):
        #创建新目录
        os.makedirs(newfiledir)
    shutil.copy(oldfile,newfile)

#复制文件到目标目录位置
def copyFiles(sourceDir,  targetDir):
    if sourceDir.find(".svn") > 0:
        return
    for file in os.listdir(sourceDir):
        sourceFile = os.path.join(sourceDir,  file)
        targetFile = os.path.join(targetDir,  file)
        if os.path.isfile(sourceFile):
            if not os.path.exists(targetDir):
                os.makedirs(targetDir)
            if not os.path.exists(targetFile) or(os.path.exists(targetFile) and (os.path.getsize(targetFile) != os.path.getsize(sourceFile))):
                open(targetFile, "wb").write(open(sourceFile, "rb").read())
        if os.path.isdir(sourceFile):
            First_Directory = False
            copyFiles(sourceFile, targetFile)

#判断是否是文件
def checkfileexist(filename):
    return os.path.isfile(filename)

#判断路径是否存在
def checkPathexist(dirpath):
    return os.path.isdir(dirpath)
# 读取文件内容并打印
from django.db import connection
def readFile(filename):
    fopen = open(filename, 'r', encoding='UTF-8')  # r 代表read
    with connection.cursor() as cursor:
        for eachLine in fopen:
            cursor.execute(eachLine)
    print("读取到得内容如下：", eachLine)
    fopen.close()


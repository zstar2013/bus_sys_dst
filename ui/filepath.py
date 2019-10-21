import os
import sys



#返回当前路径目录
def get_local_dir(path):
    dirs=os.listdir(path)
    return dirs

#判断是否是文件
def checkfileexist(filename):
    return os.path.isfile(filename)

if __name__ == "__main__":
    s=get_local_dir("G://")
    print(s)
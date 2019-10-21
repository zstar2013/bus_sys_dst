
import re
zh_pattern = re.compile(u'[\u4e00-\u9fa5]+')

def containsAny(allstr, childstr):
    for item in filter(childstr.__contains__, allstr):  # python3里直接使用filter
        return True
    return False

def containsAnyOr(allstr, childstrlist):
    for childstr in childstrlist:
        for item in filter(childstr.__contains__, allstr):  # python3里直接使用filter
            return True
    return False

def contains(allstr, childstr):
    return childstr in allstr

def contain_zh(word):
    '''
    判断传入字符串是否包含中文
    :param word: 待判断字符串
    :return: True:包含中文  False:不包含中文
    '''
    global zh_pattern
    match = zh_pattern.search(word)

    return match

def lowercase(string):
    return string.lower()

def uppercase(string):
    return string.upper()

def isContainOr(keys,target):
    for key in keys:
        if isContain(key,target):
            return True
    return False

def isContain(key,target):
    return contains(target, key)
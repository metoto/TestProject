# -*- coding: utf-8 -*-

import string
import re


class ExtractorLogfile:

    #sample = u'正则表达式是一种很有用的处理文本的工具。'

    #usample = unicode(sample, 'utf8')



    @staticmethod
    def findPart(regex, text, name):
        res = re.findall(regex, text)
        print "There are %d %s parts:" % (len(res), name)
        for r in res:
            print r



    @staticmethod
    def loadFile(path):

        f = open(path)
        return f

    @staticmethod
    def readLines(file):
        str_f = file.readlines()
        print(str_f[2])
        return str_f

    @staticmethod
    def countKeywords(keyword, str):
        print str
        print keyword
        print unicode(str,'utf8')
        list = re.findall(keyword,unicode(str, 'utf8'))
        #unicode(str, 'utf8')
        print list
        pass



if __name__ == "__main__":
    f = ExtractorLogfile.loadFile(u'C:\\Users\\Julia\\PycharmProjects\\TestProject\\assert\\17_qiche.txt')
    str = ExtractorLogfile.readLines(f)
    #print str
    count= 0
    for line in str:
        usample = unicode(line,'utf8')
        res = re.findall(u"\"\u6c7d\u8f66\"200", usample)
        if len(res)>0:
            count+=1

    print count
        #ExtractorLogfile.findPart(u"\"\u5bb6\u5c45\"200", usample, "unicode chinese")

    #usample = unicode(str, 'utf8')


    #ExtractorLogfile.findPart(u"\"\u5bb6\u5c45\"200", usample, "unicode chinese")
    pass



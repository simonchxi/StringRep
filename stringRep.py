# -*- coding: utf-8 -*-
"""
@time: 2021-11-08
@author: simon


""" 
import docx

# 传入文件(file),将旧内容(old_content)替换为新内容(new_content)
def replace(infile, outfile, old_content, new_content):
    content = read_file(infile)

    content = content.replace(old_content, new_content)
    








    rewrite_file(outfile, content)

# 读文件内容
def read_file(infile):
    with open(infile, encoding='UTF-8') as f:
        read_all = f.read()
        f.close()
        print ("readfile done!")  

    return read_all

# 写内容到文件
def rewrite_file(outfile, data):
    with open(outfile, 'w', encoding='UTF-8') as f:
        f.write(data)
        f.close()

# 替换操作(将test.txt文件中的'Hello World!'替换为'Hello Qt!')
replace(r'2021 10 October-Original test file.doc', 'Oct testout.doc','ATTACHED', '公寓')


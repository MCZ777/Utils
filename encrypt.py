# -*- coding: utf-8 -*-
"""
@Time: 2024/3/13 9:25
@Auth: MCZ
@File: encrypt.py
@IDE: PyCharm
@Motto: ABC(Always Be Coding)
"""
import os
import sys
from distutils.core import setup
from Cython.Build import cythonize

# 不需要编译的文件
global exclude_list, exclude_path
exclude_list = ['encrypt.py', 'main.py']
exclude_path = ["__pycache__", ".ipynb_checkpoints"]
global pylist
pylist = []


# 功能，遍历搜索目录下所有 py 文件并返回列表
def search(basedir, target, exclude):
    # 主目录下的所有文件，文件夹集合
    items = os.listdir(basedir)
    print("items: ", items)
    for item in items:
        # 拼接
        path = os.path.join(basedir, item)
        # path 是目录，继续搜索
        if os.path.isdir(path) and path.split('/')[-1] not in exclude_path:
            print('[-]', path)
            search(path, target, exclude)
        # 不是目录，取最后一列值,判断是否以 target 结尾,并且为非排除文件
        elif path.split('/')[-1].endswith(target) and path.split('/')[-1] not in exclude:
            print('[+]', path)
            pylist.append(str(path))
        else:
            pass
            # print('[!]',path)
    print(pylist)
    return pylist


# 根据输入构建 setup.py 文件
def newSetupFile(exclude_list, filename):
    str = '''
import os
import sys
from distutils.core import setup
from Cython.Build import cythonize


exclude_list = {} # 不需要编译的py文件
if __name__ == '__main__':
    setup(ext_modules = cythonize('{}',exclude=exclude_list))


    '''.format(exclude_list, filename)
    file = open('setup.py', "w")
    file.write(str)
    file.close


# 清理setup 构建生成临时文件
def cleanSetupFile(filepath, filename):
    os.chdir(filepath)
    print('======{}/{}'.format(filepath, filename))
    os.system("rm -rf build")
    os.remove('setup.py')
    os.remove(filename.split('.')[0] + '.c')
    os.remove(filename)


def main():
    # list = search('./', '.py', exclude_list)

    list = search(os.getcwd(), '.py', exclude_list)
    for filepath in list:
        os.path.split(filepath)
        filename = os.path.split(filepath)[1]
        filepath = os.path.split(filepath)[0]
        # 切换目录
        os.chdir(filepath)
        # 写入setup.py 文件
        newSetupFile(exclude_list, filename)
        os.system("python setup.py build_ext --inplace")
        cleanSetupFile(filepath, filename)


if __name__ == '__main__':
    print('\033[1;31;40m 加密py文件为{} 路径下所有py文件，如有排除请更新exclude_list 列表 \033[0m'.format(os.getcwd()))

    tag = input('请确认: \033[1;31;40m 加密后会删除源文件，请确认执行前是否备份: yes/no \033[0m')
    if tag == 'yes':
        main()
    else:
        print('\033[1;31;40m 请备份后执行加密操作 \033[0m')

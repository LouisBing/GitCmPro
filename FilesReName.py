#-*-coding:utf-8-*-

import os, time
from pandas import Series, DataFrame, np, ExcelWriter
import pandas as pd
import TxtOperator

def fileRename(df):
    dir = df['目录名']
    oldFile = df['文件名']
    newFile = df['重命名']
    Olddir = os.path.join(dir, oldFile)
    Newdir = os.path.join(dir, newFile)
    os.rename(Olddir, Newdir)

def getFileMtime(df):
    dir = df['目录名']
    oldFile = df['文件名']
    # newFile = df['重命名']
    Olddir = os.path.join(dir, oldFile)
    t = os.path.getmtime(Olddir)
    mt = time.strftime("%Y%m%d%H%M%S", time.localtime(t))
    return mt

def creatRefile_xls(dir,refile_xls):
    dflist = []
    for root, dirs, files in os.walk(dir, topdown=True):
        tempDf = pd.DataFrame(columns=['目录名', '文件名'])
        filesAndDirs = files+dirs
        # 以下错误代码:因此遍历中files、dirs都为不可变量，extend会修改原始列表，导致filesAndDirs为空。
        # filesAndDirs = files.extend(dirs)
        tempDf['文件名'] = filesAndDirs
        tempDf['目录名'] = root
        dflist.append(tempDf)

    df = pd.concat(dflist)
    # print(df)


    # l = os.listdir(dir)
    # df = pd.DataFrame(l, columns=['文件名'])
    # df['目录名'] = dir

    df['重命名'] = np.nan
    df['更新时间'] = df.apply(getFileMtime, axis=1)
    # 创建的时候指定列的顺序
    df = df.reindex(columns=['更新时间', '目录名', '文件名', '重命名'])

    # 数据输出
    # tNow = time.strftime("%H%M%S", time.localtime())
    # 单表输出
    df.to_excel(refile_xls)

    print('配置表格已生成，请对表格进行配置。配置完成请输入：1。')

def readRefile_xls(refile_xls):
    redf = pd.read_excel(refile_xls)
    redf = redf[redf['重命名'].notna()]
    redf.apply(fileRename, axis=1)

# dir = 'F:\个人文件夹\Pycharm附件\重命名测试'
# ####Txt读取变量值####
txtFile = u'Inputs\FilesRename.txt'
inputsList = TxtOperator.readTxt2List(txtFile,False)
dir = inputsList[0]

refile_xls = dir + '\\' + dir[dir.rfind('\\')+1:] + '_FilesRename_' + '.xlsx'
# refile_xls = r'F:\个人文件夹\Pycharm附件\重命名测试\重命名测试_FilesRename-202200.xlsx'
isx = os.path.exists(refile_xls)

if(isx):
    print('重命名配置表格已存在，请确认是否直接执行重命名.1.确认。2.重新生成配置表格。')
    next = input('请输入：')
    if(next=='1'):
        readRefile_xls(refile_xls)
    else:
        creatRefile_xls(dir, refile_xls)
        next = input('请输入：')
        if (next == '1'):
            readRefile_xls(refile_xls)
else:
    creatRefile_xls(dir, refile_xls)
    next = input('请输入：')
    if (next == '1'):
        readRefile_xls(refile_xls)
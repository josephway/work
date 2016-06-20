# coding: utf-8

import pandas as pd
import os

path = u"E:/EBSNow/产品周报/成交回报/20160613/"
files = os.listdir(path)
col_in_ex = [u'基金名称', u'委托方向', u'证券代码', u'证券名称', u'成交数量', u'成交金额', u'成交均价']
col_outex = [u'基金名称', u'委托方向', u'证券代码',
             u'证券名称', u'数量', u'金额', u'价格', u'业务过程分类']
col_NIB = [u'基金名称', u'委托方向', u'证券代码', u'证券名称', u'成交数量', u'成交金额', u'净价价格']


def in_ex(name):
    df_in = pd.read_excel(path + name, skip_footer=1).loc[:, col_in_ex]
    df_in.set_axis(1, col_in_ex)
    return df_in


def out_ex(name):
    df_out = pd.read_excel(path + name, skip_footer=1).loc[:, col_outex]
    criterion = df_out[u'业务过程分类'] == u'成交确认'
    df_out = df_out[criterion].iloc[:, 0:7]
    df_out.set_axis(1, col_in_ex)
    return df_out


def nib(name):
    df_nib = pd.read_excel(path + name, skip_footer=1).loc[:, col_NIB]
    df_nib.set_axis(1, col_in_ex)
    return df_nib

df = pd.DataFrame()

for f in files:
    if f.find(u'_成交回报') > -1:
        df = df.append(in_ex(f), ignore_index=True)
    elif f.find(u'场外') > -1:
        df = df.append(out_ex(f), ignore_index=True)
    elif f.find(u'银行间') > -1:
        df = df.append(nib(f), ignore_index=True)

df = df.set_index([u'基金名称', u'证券代码']).sort_index()

df.to_excel("c:/work/records0613.xls")

# -*- coding: utf-8 -*-
import pandas as pd
import os
path = 'c:/work/bonds/'
files = map(lambda x: x.decode('GBK'), os.listdir(path))
cols = [u'日期', u'基金名称', u'证券代码', u'证券名称', u'持仓', u'成本价']
df = pd.read_excel(path + files[0], skip_footer=1).loc[:, cols]
for i in files[1:]:
    df = df.append(pd.read_excel(path + i, skip_footer=1).loc[:, cols])
df.set_index([u'日期', u'基金名称', u'证券代码'], inplace='true')
df.sort_index(inplace=True)
df = df[(df[u'持仓'] > 0) & (df[u'成本价'] > 0)]
df[u'成本价'] = df[u'成本价'].round(4)
df[u'金额'] = df[u'持仓'] * df[u'成本价']
df.to_excel('c:/work/' + "all_bonds.xls")

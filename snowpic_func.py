import requests as req
import json
import time
import random
from lxml import etree
import pandas as pd
import random
import re
from pandas import *
import numpy as np
import dataframe_image as dfi
import os

# 数据分析函数
    
# 数据清洗
def get_CareValue(care_list):
    care_float = []
    for care in care_list:
        care = str(care)
        care = re.sub("\s|球友关注","",care) #清洗
        if re.search("万",care):  #转换为数字
            care_num = int(float(care[:-1])*10000)
        elif re.search(r"\d",care) and not re.search("万",care):   #'''如果未报告球友关注'''
            care_num = int(float(care))
        else:
            care_num = 0
        care_float.append(care_num)
    return care_float

def get_FluctValue(fluct_list):
    fluct_float = []
    for fluct in fluct_list:
        fluct = str(fluct)
        if re.search("%",fluct):
            fluct_temp = re.search("(?<=\s)([\d\D]*?)(?=%)",fluct,re.M|re.I).group(1)
            fluct_num = float(fluct_temp)
        elif re.search("\d",fluct) and not re.search("%",fluct):
            fluct_num = float(fluct)
        else: #ST未报告涨跌幅
            fluct_num = 0
        fluct_float.append(fluct_num)
    return fluct_float

def get_PriceValue(price_list):
    price_float = [float(re.sub(r'''H|K|\$|¥''','',price)) for price in price_list] 
    return price_float

#################################################################################
##                                 重新输出沪深趋势表             
#################################################################################

def get_analysis_HS(time_1,time_0,weekday):
    # 进行数据分析,返回趋势表
    # 前一日行情
    df_0 = pd.read_excel(r'..\雪球20210228\雪球网沪深证券关注度.xlsx',sheet_name = "{}".format(time_0)) 
    if not os.path.exists(r".\雪球网_沪深_关注人数_趋势.xlsx"):
        df_new = pd.read_excel(r'..\雪球20210228\雪球网沪深证券关注度.xlsx',sheet_name = "{}".format(time_1),usecols = ['证券代码','证券名称']) 
        with ExcelWriter(r".\雪球网_沪深_关注人数_趋势.xlsx",mode='w',encoding='utf-8',engine='openpyxl') as writer:
            df_new.to_excel(writer,index = False)
    
    care_0 = df_0["关注人数"].values.tolist()
    care_0 = get_CareValue(care_0)
    fluc_0 = df_0["涨跌幅"].tolist()
    fluc_0 = get_FluctValue(fluc_0)
    # ClosePrice_0 = df_0["收盘价"].values.tolist()
    # ClosePrice_0 = get_PriceValue(ClosePrice_0)
    # 当日行情
    df_1 = pd.read_excel(r'..\雪球20210228\雪球网沪深证券关注度.xlsx',sheet_name = "{}".format(time_1)) 
    care_1 = df_1["关注人数"].tolist()
    care_1 = get_CareValue(care_1)
    fluc_1 = df_1["涨跌幅"].tolist()
    fluc_1 = get_FluctValue(fluc_1)
    ClosePrice_1 = df_1["收盘价"].values.tolist()
    ClosePrice_1 = get_PriceValue(ClosePrice_1)
    # 行情变动
    care_d =  [c1 - c0 for c1,c0 in zip(care_1,care_0)]
    fluc_d = [f1 - f0 for f1,f0 in zip(fluc_1,fluc_0)]
    care_dp = []
    for c0,cd in zip(care_0,care_d):
        if c0 == 0:
            dp = 'Null'
        else:
            dp = cd/c0
        care_dp.append(dp)

    df_trend_0 = pd.read_excel(r".\雪球网_沪深_关注人数_趋势.xlsx") #读取当日之前的关注趋势 
    df_trend_1 = pd.DataFrame([care_1,care_d,care_dp,fluc_d,fluc_1,ClosePrice_1]).T #当日的关注变动
    df_trend_1.columns = ["关注人数_{}".format(time_1),"关注增量_{}".format(time_1),"关注增幅_{}".format(time_1),"涨跌幅变动(%)_{}".format(time_1),"涨跌幅_{}".format(time_1),"收盘价_{}".format(time_1)]
    df_trend_1['周几_{}'.format(time_1)] = [weekday]*len(ClosePrice_1)
    df_trend = pd.concat([df_trend_0,df_trend_1],axis=1) #拼接
    df_trend.to_excel(".//雪球网_沪深_关注人数_趋势.xlsx",index=False) #截止当日的关注趋势
    # return df_trend
    print("已追加输出趋势表。")

# for i in range(31):
#     time_1 = time.strftime("%Y-%m-%d",time.localtime(time.time()-86400*i)) #当日
#     time_0 = time.strftime("%Y-%m-%d",time.localtime(time.time()-86400*(i+1))) #前一日
#     weekday = time.strftime("%w",time.localtime(time.time()-86400*i))  #今天周几
#     print(time_1)
#     get_analysis_HS(time_1,time_0,weekday)


##################################################################################
#                                   重新输出港股趋势表
##################################################################################
def get_analysis_HK(time_1,time_0,weekday):
    # 进行数据分析,返回趋势表
    # 前一日行情
    df_0 = pd.read_excel(r'..\雪球20210228\雪球网HK证券关注度.xlsx',sheet_name = "{}".format(time_0)) 
    if not os.path.exists(r".\雪球网_HK_关注人数_趋势.xlsx"):
        df_new = pd.read_excel(r'..\雪球20210228\雪球网HK证券关注度.xlsx',sheet_name = "{}".format(time_1),usecols = ['证券代码','证券名称']) 
        with ExcelWriter(r".\雪球网_HK_关注人数_趋势.xlsx",mode='w',encoding='utf-8',engine='openpyxl') as writer:
            df_new.to_excel(writer,index = False)
    
    care_0 = df_0["关注人数"].values.tolist()
    care_0 = get_CareValue(care_0)
    fluc_0 = df_0["涨跌幅"].tolist()
    fluc_0 = get_FluctValue(fluc_0)
    # ClosePrice_0 = df_0["收盘价"].values.tolist()
    # ClosePrice_0 = get_PriceValue(ClosePrice_0)
    # 当日行情
    df_1 = pd.read_excel(r'..\雪球20210228\雪球网HK证券关注度.xlsx',sheet_name = "{}".format(time_1)) 
    care_1 = df_1["关注人数"].tolist()
    care_1 = get_CareValue(care_1)
    fluc_1 = df_1["涨跌幅"].tolist()
    fluc_1 = get_FluctValue(fluc_1)
    ClosePrice_1 = df_1["收盘价"].values.tolist()
    ClosePrice_1 = get_PriceValue(ClosePrice_1)
    # 行情变动
    care_d =  [c1 - c0 for c1,c0 in zip(care_1,care_0)]
    fluc_d = [f1 - f0 for f1,f0 in zip(fluc_1,fluc_0)]
    care_dp = []
    for c0,cd in zip(care_0,care_d):
        if c0 == 0:
            dp = 'Null'
        else:
            dp = cd/c0
        care_dp.append(dp)

    df_trend_0 = pd.read_excel(r".\雪球网_HK_关注人数_趋势.xlsx") #读取当日之前的关注趋势 
    df_trend_1 = pd.DataFrame([care_1,care_d,care_dp,fluc_d,fluc_1,ClosePrice_1]).T #当日的关注变动
    df_trend_1.columns = ["关注人数_{}".format(time_1),"关注增量_{}".format(time_1),"关注增幅_{}".format(time_1),"涨跌幅变动(%)_{}".format(time_1),"涨跌幅_{}".format(time_1),"收盘价_{}".format(time_1)]
    df_trend_1['周几_{}'.format(time_1)] = [weekday]*len(ClosePrice_1)
    df_trend = pd.concat([df_trend_0,df_trend_1],axis=1) #拼接
    df_trend.to_excel(".//雪球网_HK_关注人数_趋势.xlsx",index=False) #截止当日的关注趋势
    # return df_trend
    print("已追加输出HK趋势表。")
    
# for i in range(1,31):
#     time_1 = time.strftime("%Y-%m-%d",time.localtime(time.time()-86400*i)) #当日
#     time_0 = time.strftime("%Y-%m-%d",time.localtime(time.time()-86400*(i+1))) #前一日
#     weekday = time.strftime("%w",time.localtime(time.time()-86400*i))  #今天周几
#     print(time_1,weekday)
#     get_analysis_HK(time_1,time_0,weekday)

##################################################################################
#                                    作图函数
##################################################################################
import dataframe_image as dfi
import seaborn as sns
from matplotlib import pyplot as plt
import random

def get_pic(Market, QueryList):
    #作图函数3
    # 外嵌图表plt.table()
     '''
     StkName:证券名称，一个列表
     '''
     def keep_stk(df):
        df2 = pd.DataFrame()
        NamesList = df["证券名称"].tolist()
        KeepList = [name for name in NamesList for Query in QueryList if re.search(Query,name)]
        for keep in KeepList:
            df1 = df[df["证券名称"].isin([keep])]
            df2 = df2.append(df1,ignore_index=True)
        return df2
     # 保留待查询的证券
     if Market == 'HS':
         df = pd.read_excel(".//雪球网_沪深_关注人数_趋势.xlsx")
     else:
         df = pd.read_excel(".//雪球网_HK_关注人数_趋势.xlsx")
     columns = df.columns.tolist()
     columns = [re.sub(r'-|\.','',co) for co in columns]
     df.columns = columns
     wide = keep_stk(df) 
     # j的值只能是数字，不能含有任何字符串
     long = pd.wide_to_long(wide,
        stubnames=['关注人数','关注增量','关注增幅','涨跌幅变动(%)','涨跌幅','收盘价','周几'],
        i = '证券名称',
        j = '日期',
        sep = '_'
        )
     long = long.sort_values(by=['证券名称','日期'])
     long = long[long['周几'].isin(['1','2','3','4','5'])]
     # df_image = long.drop(columns=['Unnamed: 0'])
     # df_image.to_excel(r'查询证券面板趋势.xlsx')
     # chrome_path = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
     # dfi.export(df_image, ".//查询证券趋势_{}至今.png".format(time_1), fontsize=14,max_rows=60, max_cols=10, table_conversion='chrome',chrome_path=chrome_path)
     '''
     作图：
     散点图：x轴关注人数，y轴收盘价
     '''
     grouped = long.groupby(by=['证券名称'])
     # colors=['blue', 'orange', 'red', 'green', 'purple']
     p=0
     for stk,group in grouped:
         stkname = re.search(r'(.*?)(?=（|\()', stk).group(0)
         print(stkname)
              # 设置正常显示中文标签
         plt.rcParams['font.sans-serif'] = ['SimHei']
        # 正常显示负号
         plt.rcParams['axes.unicode_minus'] = False
        # 设置字体大小
         plt.rcParams.update({'font.size':12})
             # 组合图1
         plt.figure(figsize=(14,8))
         plt.subplot(111)
         # sns.palplot(sns.color_palette('hls', 1))
         # plt.show()  
        
         x = group['关注人数']
         y= group['收盘价']
         plt.scatter(x, y,s=20, c='r', linewidths=3)
         # plt.xticks(x,group.index.get_level_values(level=1),rotation=80)
         plt.xlabel('关注人数')
         plt.ylabel('收盘价')
         # plt.ylim(0,max(y)+200)
    
         plt.title(QueryList[p], fontsize='14', color='k')
         plt.savefig("%s关注趋势.png" %stkname,bbox_inches = 'tight')
         plt.show()
         p+=1
         print("已输出图片%s/%s" %(p,3))

# -*- coding: utf-8 -*-
"""
Created on Thu Feb 18 18:40:31 2021

@author: dell
"""

################先设函数################
# 第一个页面，获取证券代码函数
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

def get_stkcds():
    print("此命令用于更新获取证券代码，一月更新一次，更新当日不做数据分析，更新后进行数据分析前应该创建新的趋势文件excel")
    symbols = []
    headers = {
        'Cookie': "acw_tc=2760820016133179103925414ecdc10f58757af33c495ac713442e0d55d8b2; xq_a_token=62effc1d6e7ddef281d52c4ea32f6800ce2c7473; xqat=62effc1d6e7ddef281d52c4ea32f6800ce2c7473; xq_r_token=53a0f79d5bae795fb7abc6814dc0fc0410413016; xq_id_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJ1aWQiOi0xLCJpc3MiOiJ1YyIsImV4cCI6MTYxNTYwMzIxNSwiY3RtIjoxNjEzMzE3OTA0NzA0LCJjaWQiOiJkOWQwbjRBWnVwIn0.p965OyxmqoTfT2EWM-RMiPEW7TyEYClA9gZ6PaqE6YsRaficE9jAnMSGHVgZq3K7OpJJGhD6SDbQQ0OaujXufZ2V1jaNYklbpuo82qw4aPxSS3q0H9EMA6XdwGIcIbCs4Da4qAVuChE58WZ3-avvRrvxS77KTm6QLQs-1fFkroVnvPYfg2xrK39aMMpxTqJlPCcpHzaO1OjNW3Pwc5Rd8tM8d1r5MSReMtFdJhJakaAtW96mp3FVha1I9U2sKmJ_NSmMyaB9cW4uM-Ixe3dnpR7QLWHuxX2_xT35TfPe4hhyw8vuYvSeQAU2zSosMgVSyzWR24CEZJyTOV5duA_ZZg; u=441613317910395; device_id=24700f9f1986800ab4fcc880530dd0ed; Hm_lvt_1db88642e346389874251b5a1eded6e3=1613317918; s=d112o4qu1u; __utma=1.148514660.1613317929.1613317929.1613317929.1; __utmc=1; __utmz=1.1613317929.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); __utmt=1; Hm_lpvt_1db88642e346389874251b5a1eded6e3=1613318258; __utmb=1.2.10.1613317929",
        'Host': 'xueqiu.com',
        'Referer': 'https://xueqiu.com/hq',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site':'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest'
    }
    for page in range(1,150):
#         print("爬取进度{}/150".format(page))
        url = "https://xueqiu.com/service/v5/stock/screener/quote/list?page={}&size=90&order_by=percent&type=sh_sz".format(page)

        i = 1
        while i < 4:
            try:
                response = req.get(url,headers = headers,timeout = 5)
                break
            except Exception as e:
                print("第{}次连接失败，重新连接".format(i))
                i += 1

        content = response.content
        content_dict = json.loads(content)
        symbol_list = content_dict.get('data').get('list')
        if symbol_list == []:
            break

        for sy in symbol_list:
            symbol = sy['symbol']
            symbols.append(symbol)

        time.sleep(random.random()*3)
    print('爬取stkcds任务结束')
    file = open(".//雪球网沪深证券代码.csv",'a',encoding = "utf-8")
    for sy in symbols:
        file.write(sy+",")
    file.close()
    print("雪球网沪深证券代码.csv已更新")

    return symbols

# 第二个页面，爬取关注情况函数
# 数据清洗
def get_CareValue(care_list):
    care_float = []
    for care in care_list:
        care = re.sub("\s|球友关注","",care) #清洗
        if re.search("万",care):  #转换为数字
            care_num = int(float(care[:-1])*10000)
        else:
            care_num = int(float(care))
        care_float.append(care_num)
    return care_float
#     print(care_float)

# 清洗涨跌幅为数值
def get_FluctValue(fluct_list):
    fluct_float = []
    for fluct in fluct_list:
        if re.search("%",fluct):
            fluct_temp = re.search("(?<=\s)([\d\D]*?)(?=%)",fluct,re.M|re.I).group(1)
            fluct_num = float(fluct_temp)
        else: #ST未报告涨跌幅
            fluct_num = 0
        fluct_float.append(fluct_num)
    return fluct_float

def get_CareNum(time_1):
    # 主函数,返回当日关注人数DataFrame df_1
    name_list = []
    price_list = []
    fluct_list = []
    care_list = []
    stkcd_list = []

    p = 1
    file = open("雪球网沪深证券代码.csv",encoding = "utf-8")
    symbols_csv = file.read()
    symbols = symbols_csv.split(",")
    symbols =[sy for sy in symbols if sy != '']
    file.close
    for sy in symbols:
        print("第二步进度为{}/4285".format(p))
        url_2 = "https://xueqiu.com/S/{}".format(sy)
        # print(url_2)
        headers = {
            "Cookie": "xq_a_token=62effc1d6e7ddef281d52c4ea32f6800ce2c7473; xqat=62effc1d6e7ddef281d52c4ea32f6800ce2c7473; xq_r_token=53a0f79d5bae795fb7abc6814dc0fc0410413016; xq_id_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJ1aWQiOi0xLCJpc3MiOiJ1YyIsImV4cCI6MTYxNTYwMzIxNSwiY3RtIjoxNjEzMzE3OTA0NzA0LCJjaWQiOiJkOWQwbjRBWnVwIn0.p965OyxmqoTfT2EWM-RMiPEW7TyEYClA9gZ6PaqE6YsRaficE9jAnMSGHVgZq3K7OpJJGhD6SDbQQ0OaujXufZ2V1jaNYklbpuo82qw4aPxSS3q0H9EMA6XdwGIcIbCs4Da4qAVuChE58WZ3-avvRrvxS77KTm6QLQs-1fFkroVnvPYfg2xrK39aMMpxTqJlPCcpHzaO1OjNW3Pwc5Rd8tM8d1r5MSReMtFdJhJakaAtW96mp3FVha1I9U2sKmJ_NSmMyaB9cW4uM-Ixe3dnpR7QLWHuxX2_xT35TfPe4hhyw8vuYvSeQAU2zSosMgVSyzWR24CEZJyTOV5duA_ZZg; u=441613317910395; device_id=24700f9f1986800ab4fcc880530dd0ed; Hm_lvt_1db88642e346389874251b5a1eded6e3=1613317918; s=d112o4qu1u; __utma=1.148514660.1613317929.1613317929.1613317929.1; __utmc=1; __utmz=1.1613317929.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); Hm_lpvt_1db88642e346389874251b5a1eded6e3=1613318368; acw_tc=2760820916133197637453679e4ece6157054e00c1ccc8a52b089d7c25b25d",
            "Host": "xueqiu.com",
            "Referer": "https://xueqiu.com/hq",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36"
        }

        i = 1
        while i < 4:
            try:
                res = req.get(url_2,headers = headers, timeout = 5)
                break
            except Exception as e:
                print("第{}次连接失败，重新连接".format(i))
                i += 1

        content = res.content.decode()
        tree = etree.HTML(content)
        name_path = "//div[@class='container-sm float-left stock__main']/div[@class='stock-name']/text()"
        price_path = "//div[@class='stock-current']/strong/text()"
        fluct_path = "//div[@class='stock-change']/text()"
        care_path = "//div[@class='stock-time']/div[1]"

        name = tree.xpath(name_path)
        price = tree.xpath(price_path)
        fluct = tree.xpath(fluct_path)
        care_num = tree.xpath(care_path)

        care_num = [care.text for care in care_num]

        name_list.extend(name)
        price_list.extend(price)
        fluct_list.extend(fluct)
        care_list.extend(care_num)
        stkcd_list.append(sy)

        time.sleep(random.random()*3)
        p += 1
    # 清洗数据
    fluct_float = get_FluctValue(fluct_list)
    care_float = get_CareValue(care_list)
    
    df_1 = pd.DataFrame([stkcd_list,name_list,price_list,fluct_float,care_float]).T
    df_1.columns = ['证券代码','证券名称','收盘价','涨跌幅','关注人数']
    df_1.index = df_1.index + 1
    with ExcelWriter(".//雪球网沪深证券关注度.xlsx",mode="a",encoding="utf-8",engine="openpyxl") as writer:
        df_1.to_excel(writer,index=False,sheet_name = "{}".format(time_1))
    print("今日已追加输出关注人数。") 
    return df_1

# 数据分析函数
    
# 关注人数增量、增幅排行
# 价格涨跌幅排行

def get_analysis(time_1,time_0):
    # 进行数据分析,无返回值，直接输出追加趋势文件
    # 前一日行情
    df_0 = pd.read_excel('.//雪球网沪深证券关注度.xlsx',encoding = "utf-8",sheet_name = "{}".format(time_0)) 
    care_0 = df_0["关注人数"].tolist()
    fluc_0 = df_0["涨跌幅"].tolist()
    # 当日行情
    df_1 = pd.read_excel('.//雪球网沪深证券关注度.xlsx',encoding = "utf-8",sheet_name = "{}".format(time_1)) 
    care_1 = df_1["关注人数"].tolist()
    fluc_1 = df_1["涨跌幅"].tolist()
    # 行情变动
    care_d =  [c1 - c0 for c1,c0 in zip(care_1,care_0)]
    fluc_d = [f1 - f0 for f1,f0 in zip(fluc_1,fluc_0)]
    care_dp = [cd/c0 for cd,c0 in zip(care_d,care_0)]

    df_trend_0 = pd.read_excel(".//雪球网_沪深_关注人数_趋势.xlsx") #读取当日之前的关注趋势 
    df_trend_1 = pd.DataFrame([care_1,care_d,care_dp,fluc_d,fluc_1]).T #当日的关注变动
    df_trend_1.columns = ["关注人数_{}".format(time_1),"关注增量_{}".format(time_1),"关注增幅_{}".format(time_1),"涨跌幅变动(%)_{}".format(time_1),"涨跌幅_{}".format(time_1)]

    df_trend = pd.concat([df_trend_0,df_trend_1],axis=1) #拼接
    df_trend.to_excel(".//雪球网_沪深_关注人数_趋势.xlsx",index=False) #截止当日的关注趋势
    print("今日关注变动已追加输出。")
# with ExcelWriter(".//雪球网_沪深_关注人数_趋势.xlsx",mode="a",encoding="utf-8",engine="openpyxl") as writer:
#     df_trend.to_excel(writer,index=False,sheet_name = "sheet1")
# # 默认为xlsxwriter，要更换为openpyxl,否则会报错：Append mode is not supported with xlsxwriter
# ExcelWriter只能在同一个WorkBook下追加多个Sheet


# 作图函数1
# pip install dataframe_image
# 国内镜像pip install dataframe_image  -i http://pypi.douban.com/simple/ --trusted-host pypi.douban.com
# import dataframe_image as dfi
# dfi.export(obj, filename, fontsize=14,max_rows=None, max_cols=None, table_conversion='chrome', chrome_path=None)
    
    
import dataframe_image as dfi
# import matplotlib as plt
def get_figure_1(time_1):
    
    def drop_NST(df):
        # 去掉N和ST
        drop_list = df["证券名称"].tolist()
        drop_list = [drop for drop in drop_list if re.search("N|ST|C|B",drop)]
        for drop in drop_list:
            df = df[~df["证券名称"].isin([drop])]
        return df
    
    # 读取并清洗数据
    df_trend = pd.read_excel(".//雪球网_沪深_关注人数_趋势.xlsx",usecols=['证券代码','证券名称',"关注增量_{}".format(time_1),"关注人数_{}".format(time_1),"关注增幅_{}".format(time_1),"涨跌幅_{}".format(time_1),"涨跌幅变动(%)_{}".format(time_1)])
    
    df_trend_draw = drop_NST(df_trend) # 去掉N和ST
    # 保留关注人数增量top50
    # 从高到低排序
    df_trend_sort=df_trend_draw.sort_values(by=["关注增量_{}".format(time_1),"关注增幅_{}".format(time_1),"关注人数_{}".format(time_1),"涨跌幅_{}".format(time_1),"涨跌幅变动(%)_{}".format(time_1)],ascending = False)
    df_CareTop5 = df_trend_sort.iloc[0:10,:] #保留top50
    # 提取数据

    x_0 = df_CareTop5["证券代码"] #"关注增量"top5的公司代码
    x =  df_CareTop5["证券名称"] #"关注增量"top5的公司名称
    y1 = df_CareTop5['关注人数_{}'.format(time_1)] #关注人数
    y2 = df_CareTop5['关注增量_{}'.format(time_1)] #关注增量
    y3 = df_CareTop5["关注增幅_{}".format(time_1)] #关注增幅
    y4 = df_CareTop5["涨跌幅_{}".format(time_1)] #涨跌幅
    y5 = df_CareTop5["涨跌幅变动(%)_{}".format(time_1)] #涨跌幅变动（%）
    
    
    """
    制表:关注人数，关注增量，关注增幅，涨跌幅，涨跌幅变动
    """
    # 表格
    x0_t = x_0.tolist()  #必须使用tolist,否则报错
    x_t = [re.match(r"(.*?)(?=\()",c).group(1) for c in x.values.tolist()]
    y1_t = y1.values.tolist()
    y2_t = y2.values.tolist()
    y3_t = [str(round(y*100,2)) +"%" for y in y3.values.tolist()]
    y4_t = [str(round(y,2))+"%" for y in y4.values.tolist()]
    y5_t = [str(round(y,2)) + "个百分点" for y in y5.values.tolist()]
    
    data = [x0_t,x_t,y2_t,y3_t,y1_t,y4_t]
    rows = [i for i in range(1,11)]
    columns =  ['证券代码','证券名称','关注增量','关注增幅','关注人数','涨跌幅']
    df_image = pd.DataFrame(data).T
    df_image.columns = columns
    df_image.index = rows
    '''
    制表:关注人数，关注增量，关注增幅，涨跌幅，涨跌幅变动
    直接将df保存为图片到本地
    '''
    #     # 设置正常显示中文标签
    # plt.rcParams['font.sans-serif'] = ['SimHei']
    #     # 正常显示负号
    # plt.rcParams['axes.unicode_minus'] = False
    #     # 设置字体大小
    # plt.rcParams.update({'font.size':14})
    chrome_path = r'C:\Users\dell\AppData\Local\Google\Chrome\Application\chrome.exe'
    dfi.export(df_image, "每日沪深证券关注情况top50_{}.png".format(time_1), fontsize=14,max_rows=60, max_cols=10, table_conversion='chrome',chrome_path=chrome_path)
    print("图片top50已保存。")
'''
obj：表示的是待保存的DataFrame数据框；
filename：表示的是图片保存的本地路径；
fontsize：表示的是待保存图片中字体大小，默认是14；
max_rows：表示的是DataFrame输出的最大行数。
这个数字被传递给DataFrame的to_html方法。为防止意外创建具有大量行的图像，具有100行以上的DataFrame将引发错误。显式设置此参数以覆盖此错误，对所有行使用-1。
max_cols：表示的是DataFrame输出的最大列数。
这个数字被传递给DataFrame的to_html方法。为防止意外创建具有大量列的图像，包含30列以上的DataFrame将引发错误。显式设置此参数以覆盖此错误，对所有列使用-1。
table_conversion：‘chrome’或’matplotlib’，默认为’chrome’。DataFrames将通过Chrome或matplotlib转换为png。
除非无法正常使用，否则请使用chrome。 matplotlib提供了一个不错的选择。
可以看到：这个方法其实就是通过chrome浏览器，将这个DataFrames转换为png或jpg格式。
'''
    
# 作图函数2
def get_figure_2(time_1):
    
    def drop_NST(df):
        # 去掉N和ST
        drop_list = df["证券名称"].tolist()
        drop_list = [drop for drop in drop_list if re.search("N|ST|C|B",drop)]
        for drop in drop_list:
            df = df[~df["证券名称"].isin([drop])]
        return df
    
    # 读取并清洗数据
    df_trend = pd.read_excel(".//雪球网_沪深_关注人数_趋势.xlsx",usecols=['证券代码','证券名称',"关注增量_{}".format(time_1),"关注人数_{}".format(time_1),"关注增幅_{}".format(time_1),"涨跌幅_{}".format(time_1),"涨跌幅变动(%)_{}".format(time_1)])
    
    df_trend_draw = drop_NST(df_trend) # 去掉N和ST
    # 保留关注人数增量top50
    # 从高到低排序
    df_trend_sort=df_trend_draw.sort_values(by=["关注增量_{}".format(time_1),"关注人数_{}".format(time_1),"关注增幅_{}".format(time_1),"涨跌幅_{}".format(time_1),"涨跌幅变动(%)_{}".format(time_1)],ascending = True)
    df_CareTop5 = df_trend_sort.iloc[0:50,:] #保留top50
    # 提取数据
    x_0 = df_CareTop5["证券代码"] #"关注增量"top5的公司代码
    x =  df_CareTop5["证券名称"] #"关注增量"top5的公司名称
    y1 = df_CareTop5['关注人数_{}'.format(time_1)] #关注人数
    y2 = df_CareTop5['关注增量_{}'.format(time_1)] #关注增量
    y3 = df_CareTop5["关注增幅_{}".format(time_1)] #关注增幅
    y4 = df_CareTop5["涨跌幅_{}".format(time_1)] #涨跌幅
    y5 = df_CareTop5["涨跌幅变动(%)_{}".format(time_1)] #涨跌幅变动（%）
    
    
    """
    制表:关注人数，关注增量，关注增幅，涨跌幅，涨跌幅变动
    """
    # 表格
    x0_t = x_0.tolist()
    x_t = [re.match(r"(.*?)(?=\()",c).group(1) for c in x.values.tolist()]
    y1_t = y1.values.tolist()
    y2_t = y2.values.tolist()
    y3_t = [str(round(y*100,2)) +"%" for y in y3.values.tolist()]
    y4_t = [str(round(y,2))+"%" for y in y4.values.tolist()]
    y5_t = [str(round(y,2)) + "个百分点" for y in y5.values.tolist()]
    
    data = [x0_t,x_t,y1_t,y2_t,y3_t,y4_t,y5_t]
    rows = [i for i in range(-10,0)]
    rows.sort(reverse = True)
    columns =  ['证券代码','证券名称','关注人数','关注增量','关注增幅','涨跌幅','涨跌幅变动']
    df_image = pd.DataFrame(data).T
    df_image.columns = columns
    df_image.index = rows
    '''
    制表:关注人数，关注增量，关注增幅，涨跌幅，涨跌幅变动
    直接将df保存为图片到本地
    '''
    #     # 设置正常显示中文标签
    # plt.rcParams['font.sans-serif'] = ['SimHei']
    #     # 正常显示负号
    # plt.rcParams['axes.unicode_minus'] = False
    #     # 设置字体大小
    # plt.rcParams.update({'font.size':14})
    chrome_path = r'C:\Users\dell\AppData\Local\Google\Chrome\Application\chrome.exe'
    dfi.export(df_image, "每日沪深证券关注情况last50_{}.png".format(time_1), fontsize=14,max_rows=60, max_cols=10, table_conversion='chrome',chrome_path=chrome_path)
    print("图片last50已保存。")  


#作图函数3
# 外嵌图表plt.table()

def get_figure_3(time_1):
    
    def drop_NST(df):
        # 去掉N和ST
        drop_list = df["证券名称"].tolist()
        drop_list = [drop for drop in drop_list if re.search("N|ST|C",drop)]
        for drop in drop_list:
            df = df[~df["证券名称"].isin([drop])]
        return df
    
    # 读取并清洗数据
    df_trend = pd.read_excel(".//雪球网_沪深_关注人数_趋势.xlsx")
    df_trend_draw = drop_NST(df_trend) # 去掉N和ST
    # 保留关注人数增量top5
    
    df_trend_sort=df_trend_draw.sort_values(by=["关注增量_{}".format(time_1),"涨跌幅_{}".format(time_1),"关注增幅_{}".format(time_1),"涨跌幅变动(%)_{}".format(time_1)],ascending = False)
    df_CareTop5 = df_trend_sort.iloc[0:5,:] #保留top5
    df_CareTop5
    
    # 提取数据
    x =  df_CareTop5["证券名称"] #"关注增量"top5的公司名称
    y1 = df_CareTop5['关注增量_{}'.format(time_1)] #关注增量
    y2 = df_CareTop5["关注增幅_{}".format(time_1)] #关注增幅
    y3 = df_CareTop5["涨跌幅变动(%)_{}".format(time_1)] #涨跌幅变动（%）
    y4 = y3/100
    y5 = df_CareTop5['涨跌幅_{}'.format(time_1)]
    y6 = y5/100
    
    # 设置正常显示中文标签
    plt.rcParams['font.sans-serif'] = ['SimHei']
    # 正常显示负号
    plt.rcParams['axes.unicode_minus'] = False
    # 设置字体大小
    plt.rcParams.update({'font.size':12})
    
    plt.figure(figsize=(14,8))
    plt.subplot(111)
    
    # 柱形宽度
    bar_width = 0.35
    
    # 在主坐标轴绘制柱形图
    color = ['#dc5034' if i > 0 else 'g' for i in y1]
    plt.bar(x,y1,bar_width,color=color,label='关注增量')
    
    # 设置坐标轴的取值范围，避免柱子过高而与图例重叠
    plt.ylim(0,max(y1.max(),y2.max())*1.2)
    # 设置图例
    plt.legend(loc='upper left')
    
    # 设置横坐标的标签
    plt.xticks(x)
    
    # 右边折线图
    plt.twinx()
    plt.plot(x,y6,ls='-',lw=2,color='b',marker='o',label='涨跌幅')
    # 设置次坐标轴的取值范围，避免折线图波动过大
    # plt.ylim(0,1.35)
    
    # 设置图例
    plt.legend()
    
    # 图下嵌表
    y1_t = y1.values.tolist()
    y2_t = [str(round(y*100,2)) +"%" for y in y2.values.tolist()]
    y3_t = [str(round(y,2)) + "个百分点" for y in y3.values.tolist()]
    y5_t = [str(round(y,2))+"%" for y in y5.values.tolist()]
    data = [y1_t,y2_t,y5_t,y3_t]
    columns = [re.match(r"(.*?)(?=\()",c).group(1) for c in x.values.tolist()]
    rows =  ['关注增量','关注增幅','涨跌幅','涨跌幅变动']
    
    plt.table(cellText = data,
          cellLoc = 'center',
          cellColours = None,
          rowLabels = rows,
          rowColours = plt.cm.Reds(np.linspace(0, 0.5,5))[::-1],  # BuPu可替换成其他colormap
          colLabels = columns,
          colColours = plt.cm.Reds(np.linspace(0, 0.5,5))[::-1], 
          rowLoc='right',
          loc='bottom')
    #调整
    # plt.subplots_adjust(left=1.0, bottom=1.2)
       # 定义显示百分号的函数
    def to_percent(number, position=0):
        return str(round(number * 100,2)) + '%'
    
    # 次坐标轴的标签显示百分号 FuncFormatter：自定义格式函数包
    from matplotlib.ticker import FuncFormatter
    plt.gca().yaxis.set_major_formatter(FuncFormatter(to_percent))
    
    
    # 设置标题
    plt.title('\n{}沪深证券关注情况\n'.format(time_1),fontsize=36,loc='center',color = 'k')
    plt.savefig("每日沪深证券关注情况3_{}.png".format(time_1),bbox_inches = 'tight')
    # plt.show()   
    """
    一条折线展示"涨跌幅变动（%）"
    一个条形图展示"关注增量"。
    图外表格显示关注增量水平值、关注增幅水平值、涨跌幅变动水平值。
    主坐标轴是柱形图对应的数据，次坐标轴是折线图对应的数据，
    下边的横坐标轴表示"关注增量"top10的公司名称。
    """
  
    
# 自动发送微信消息函数
import os
import win32gui #pywin32-221.win-amd64-py3.7.exe
import win32con
from ctypes import *
import win32clipboard as w
import time
from PIL import Image #pip install pillow
#pip install -i https://pypi.douban.com/simple pillow
import win32api
 
#发送文字
def setText(info):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, info)
    w.CloseClipboard()
#发送图片
def setImage(imgpath):
    im = Image.open(imgpath)
    im.save('1.bmp')
    aString = windll.user32.LoadImageW(0, r"1.bmp", win32con.IMAGE_BITMAP, 0, 0, win32con.LR_LOADFROMFILE)
    
    if aString != 0:  ## 由于图片编码问题  图片载入失败的话  aString 就等于0
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_BITMAP, aString)
        w.CloseClipboard()  
# 微信搜索框不会自动获取焦点，故需要模拟鼠标点击到搜索框的位置
def m_click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
def pasteInfo():
    win32api.keybd_event(17,0,0,0)  #ctrl键位码是17
    win32api.keybd_event(86,0,0,0)  #v键位码是86
    win32api.keybd_event(86,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)
# 定位微信窗口，进行昵称备注的搜索（需点击两下才能获取到焦点）
def searchByUser(uname):
    hwnd = win32gui.FindWindow('WeChatMainWndForPC', '微信')
    setText(uname)
    m_click(100,40)
    time.sleep(0.5)
    m_click(100,40)
    pasteInfo()
    time.sleep(1)
    m_click(100,120)#搜索到之后点击
    #win32api.keybd_event(13,0,0,0)#回车
    #win32api.keybd_event(13,0,KEYEVENTF_KEYUP,0)
    #win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    #win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

def sendInfo():
    time.sleep(1)
    pasteInfo()
    time.sleep(1)
    win32api.keybd_event(18, 0, 0, 0) #Alt  
    win32api.keybd_event(83,0,0,0) #s
    win32api.keybd_event(83,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
    win32api.keybd_event(18,0,win32con.KEYEVENTF_KEYUP,0)

    # 发送完信息之后关闭窗口(跟QQ不一样，可以不关闭），接着搜索发送
def closeByUser(uname):
    hwnd = win32gui.FindWindow('WeChatMainWndForPC', '微信')
    win32api.keybd_event(18,0,0,0)  #Alt
    win32api.keybd_event(115,0,0,0) #F4
    win32api.keybd_event(115,0,KEYEVENTF_KEYUP,0)
    win32api.keybd_event(18,0,KEYEVENTF_KEYUP,0)
    
def sendMessage(user,time_1):
    print("执行程序时请将微信窗口拖至左上角,等待10s")
    time.sleep(10)
    searchByUser(user)
    setText('请查收{}证券关注情况表'.format(time_1))
    sendInfo()
    time.sleep(1)
    setImage('每日沪深证券关注情况2_{}.png'.format(time_1))
    sendInfo() 
    print("今日已发布雪球网沪深证券关注人数。")
    

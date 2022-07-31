from snow_func import *
from auto_send_mail import Auto_Send_Error_Mail

#from snow_func import * 和import snow_func含义不同
############执行程序###########
# 是否更新证券代码？
# get_stkcds()
#（更新证券代码后引创建新的趋势表）
# #趋势单列一个表
# df_trend = pd.read_excel('.//雪球网沪深证券关注度.xlsx',encoding = "utf-8",usecols = ["证券代码","证券名称"]) 
# # print(df_trend)
# # df_trend.columns = ["证券代码","证券名称","关注人数增量","关注人数增幅","涨跌幅变动"]
# # df_trend.index = df_trend.index + 1
# with ExcelWriter(".//雪球网_沪深_关注人数_趋势.xlsx",mode="a",encoding="utf-8",engine="openpyxl") as writer:
    # df_trend.to_excel(writer,index=False,sheet_name = "{}".format(time_1))

# 昨日是否节假日？若是，则需要更改time_0
# time_1 = "2021-02-19"
time_1 = time.strftime("%Y-%m-%d",time.localtime()) #当日
time_0 = time.strftime("%Y-%m-%d",time.localtime(time.time()-86400)) #前一日

# 爬取今日关注情况:向"关注情况"excel追加今日关注情况的sheet
df_1 = get_CareNum(time_1)
# 数据分析:追加"趋势"excel
get_analysis(time_1,time_0)
# 作图
time_1 = "2021-02-19"
get_figure_1(time_1)  #关注人数前10名
# get_figure_2(time_1)  #关注人数后50名
# get_figure_3(time_1) #条形图+图下嵌表
#自动发送图片消息
# 微信消息
# user = "xiaoyang"  
# sendMessage(user,time_1)
# 发送邮件
# SMTP服务器,这里使用163邮箱
mail_host = "smtp.163.com"
# 发件人邮箱
mail_sender = "wyttjcu@163.com"
# 邮箱授权码（代替账号密码）,注意这里不是邮箱密码
mail_license = "CYJTNKYHYAJLEKPZ"
# 收件人邮箱，可以为多个收件人
mail_receivers = "followers_price@163.com"
body_content = "请查收{}沪深证券关注情况".format(time_1)
mail_obj = Auto_Send_Error_Mail(mail_host, mail_sender, mail_license, mail_receivers,body_content,time_1)
mail_obj.send_email()
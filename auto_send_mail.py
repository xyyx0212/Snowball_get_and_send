# -*- coding: utf-8 -*-
"""
Created on Sat Feb 13 16:22:06 2021

@author: dell
"""


'''
自动发送邮件接口
'''

import smtplib
import email
# 负责构造文本
from email.mime.text import MIMEText
# 负责构造图片
from email.mime.image import MIMEImage
# 负责将多个对象集合起来
from email.mime.multipart import MIMEMultipart
from email.header import Header


class Auto_Send_Error_Mail:
    # 初始化该类
    def __init__(self, mail_host, mail_sender, mail_license, mail_receivers, body_content,time_1):
        self.mail_host = mail_host
        self.mail_sender = mail_sender
        self.mail_license = mail_license
        self.mail_receivers = mail_receivers
        self.body_content = body_content
        self.time_1 = time_1

    # 发送错误日志邮件
    def send_email(self):
        mm = MIMEMultipart('related')
        # 邮件主题
        subject_content = """请查收今日沪深证券关注情况图片~"""
        # 设置发送者,注意严格遵守格式,里面邮箱为发件人邮箱
        mm["From"] = self.mail_sender
        # 设置接受者,注意严格遵守格式,里面邮箱为接受者邮箱
        mm["To"] = self.mail_receivers
        # 设置邮件主题
        mm["Subject"] = Header(subject_content, 'utf-8')

        # 邮件正文内容
        # body_content = "请查收{}沪深证券关注情况".format(time_1)
        # 构造文本,参数1：正文内容，参数2：文本格式，参数3：编码方式
        message_text = MIMEText(self.body_content, "plain", "utf-8")
        # 构造图片
        imagefile_top = open(r"每日沪深证券关注情况top50_{}.png".format(self.time_1),'rb')
        message_image_top = MIMEImage(imagefile_top.read())
        imagefile_top.close()
        
        # imagefile_last = open(r"每日沪深证券关注情况last50_{}.png".format(self.time_1),'rb')
        # message_image_last = MIMEImage(imagefile_last.read())
        # imagefile_last.close()
        # 向MIMEMultipart对象中添加文本、图片对象
        mm.attach(message_text)
        mm.attach(message_image_top)
        # mm.attach(message_image_last)
        '''
        这里特别强调：
        	在服务器linux系统上，检验会更加的严格，所以采用SMTP_SSL()
        	python3.7以上，smtplib模块文件做了修改
        	需要改写为：
        		端口25，ssl端口为465
        		smtplib.SMTP_SSL(self.mail_host).connect(self.mail_host, 465)
        	版本以下：
        		smtplib.SMTP_SSL().connect(self.mail_host, 465)
        '''
        # stp = smtplib.SMTP()
        stp = smtplib.SMTP_SSL(self.mail_host)
        # 设置发件人邮箱的域名和端口，端口地址为25,ssl端口为465
        stp.connect(self.mail_host, 465)
        # set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
        stp.set_debuglevel(1)
        # 登录邮箱，传递参数1：邮箱地址，参数2：邮箱授权码
        stp.login(self.mail_sender, self.mail_license)
        # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
        stp.sendmail(self.mail_sender, self.mail_receivers, mm.as_string())
        print("邮件发送成功")
        # 关闭SMTP对象
        stp.quit()

# # SMTP服务器,这里使用163邮箱
# mail_host = "smtp.163.com"
# # 发件人邮箱
# mail_sender = "xxxxxxxxx@163.com"
# # 邮箱授权码,注意这里不是邮箱密码,如何获取邮箱授权码,请看本文最后教程
# mail_license = "BGDAEEMWDFDGBRVQ"
# # 收件人邮箱，可以为多个收件人
# mail_receivers = ["xxxxxxx@qq.com"]
#
# mm = MIMEMultipart('related')

# 邮件主题
# subject_content = """Python邮件测试"""
# # 设置发送者,注意严格遵守格式,里面邮箱为发件人邮箱
# mm["From"] = "sender_name<xxxxxx@163.com>"
# # 设置接受者,注意严格遵守格式,里面邮箱为接受者邮箱
# mm["To"] = "receiver_1_name<xxxxx@qq.com>"
# # 设置邮件主题
# mm["Subject"] = Header(subject_content, 'utf-8')



# 邮件正文内容
# body_content = """你好，这是一个测试邮件！"""
# # 构造文本,参数1：正文内容，参数2：文本格式，参数3：编码方式
# message_text = MIMEText(body_content, "plain", "utf-8")
# # 向MIMEMultipart对象中添加文本对象
# mm.attach(message_text)

# 二进制读取图片
# image_data = open('a.jpg', 'rb')
# # 设置读取获取的二进制数据
# message_image = MIMEImage(image_data.read())
# # 关闭刚才打开的文件
# image_data.close()
# # 添加图片文件到邮件信息当中去
# mm.attach(message_image)

# 构造附件
# atta = MIMEText(open('recognition.log', 'r').read(), 'base64', 'utf-8')
# # 设置附件信息
# atta["Content-Disposition"] = 'attachment; filename="sample.log"'
# # 添加附件到邮件信息当中去
# mm.attach(atta)

# 创建SMTP对象
# stp = smtplib.SMTP()
# # 设置发件人邮箱的域名和端口，端口地址为25
# stp.connect(mail_host, 25)
# # set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
# stp.set_debuglevel(1)
# # 登录邮箱，传递参数1：邮箱地址，参数2：邮箱授权码
# stp.login(mail_sender, mail_license)
# # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
# stp.sendmail(mail_sender, mail_receivers, mm.as_string())
# print("邮件发送成功")
# # 关闭SMTP对象
# stp.quit()
        

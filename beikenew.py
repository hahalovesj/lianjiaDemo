#!/usr/bin/python3
# -*- coding:utf-8 -*-
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from selenium import webdriver  # 导入webdriver包
import time
import xlrd

start = time.clock()  # 程序开始时间
# 打开文件
workbook = xlrd.open_workbook(r'/Volumes/sd/workSpace/vsCodeWorkplace/LianJiaDemo/目标盘.xls')
# 根据sheet索引或者名称获取sheet内容
bj_list2 = workbook.sheet_by_name('北京')
# 北京sheet的第一列内容
bj_list = bj_list2.col_values(0)
print(bj_list)

sh_list2 = workbook.sheet_by_name('上海')
# 上海sheet的第一列内容
sh_list = sh_list2.col_values(0)
print(sh_list)

driver = webdriver.Chrome()
driver.maximize_window()  # 最大化浏览器

time.sleep(1)  # 暂停1秒钟
print("---------------------【开始爬取北京楼盘数据】----------------------")
for i in range(len(bj_list)):
    driver.get("https://bj.ke.com/ershoufang/rs" + bj_list[i])  # 通过get()方法，打开一个url站点
    time.sleep(4)  # 暂停2秒钟

    try:
        ele = driver.find_element_by_xpath("//*[@id='sem_card']/div")
        if len(ele) > 0:
            chushou = driver.find_element_by_xpath("//*[@id='sem_card']/div/div[2]/div[2]/div[2]/div[2]")
            chengjiao = driver.find_element_by_xpath("//*[@id='sem_card']/div/div[2]/div[2]/div[3]/div[2]")
            daikan = driver.find_element_by_xpath("//*[@id='sem_card']/div/div[2]/div[2]/div[4]/div[2]")
            print(bj_list[i])
            print(" 正在出售" + chushou.text + " 近90天成交" + chengjiao.text + " 近30天带看" + daikan.text)
            time.sleep(2)  # 暂停2秒钟
    except:
        print("没有楼盘卡片")

    # 房源爬虫
    lists = driver.find_elements_by_xpath("//*[@id='beike']/div[1]/div[4]/div[1]/div[4]/ul//li/a")  # 找到a链接
    length = len(lists)  # 列表长度
    count = driver.find_element_by_xpath("//*[@id='beike']/div[1]/div[4]/div[1]/div[2]/div[1]/h2/span")  # 共找到x套
    if 0 < int(count.text) < 32:  # 链接超过1个小于32个
        for j in range(0, len(lists)):  # 遍历列表的循环，使程序可以逐一点击
            if j == 5:  # 第5个推广引流的a跳过
                continue
            # 在每次循环内都重新获取a标签，组成列表
            links = driver.find_elements_by_xpath("//*[@id='beike']/div[1]/div[4]/div[1]/div[4]/ul//li/a")

            link = links[j]  # 逐一将列表里的a标签赋给link
            url = link.get_attribute('href')  # 提取a标签内的链接，注意这里提取出来的链接是字符串
            print(url)
            driver.get(url)  # url为字符串，直接用浏览器打开网址即可
            time.sleep(1)  # 留出加载时间
            total = driver.find_element_by_xpath(
                "//*[@id='beike']/div[1]/div[4]/div[1]/div[2]/div[2]/span[1]").text  # .text的意思是指输出这里的纯文本内容
            print("总价" + total + "万")
            unitPrice = driver.find_element_by_xpath(
                "//*[@id='beike']/div[1]/div[4]/div[1]/div[2]/div[2]/div[1]/div[1]/span").text
            print("均价" + unitPrice + "元/平米")
            # 房屋户型
            houseType = driver.find_element_by_xpath("//*[@id='introduction']/div/div/div[1]/div[2]/ul/li[1]").text
            print(houseType)
            # 挂牌时间
            onTime = driver.find_element_by_xpath("//*[@id = 'introduction']/div/div/div[2]/div[2]/ul/li[1]").text
            print(onTime)
            driver.back()  # 后退，返回原始页面目录页
            time.sleep(1)  # 留出加载时间

    elif int(count.text) >= 32:  # 分页处理
        pages = driver.find_elements_by_xpath("//*[@id='beike']/div[1]/div[4]/div[1]/div[5]/div[2]/div/a")
        pageNum = len(pages)  # 翻页页数
        # print(pageNum)
        for i in range(pageNum):
            pages = driver.find_elements_by_xpath("//*[@id='beike']/div[1]/div[4]/div[1]/div[5]/div[2]/div/a")
            time.sleep(1)
            pages[i].click()  # 点击翻页
            time.sleep(1)
            lists = driver.find_elements_by_xpath("//*[@id='beike']/div[1]/div[4]/div[1]/div[4]/ul//li/a")  # 获得所有房子链接列表
            print("第" + str(i) + "页")
            print(len(lists))  # 列表长度

            for j in range(len(lists)):  # 遍历列表的循环，使程序可以逐一点击
                if j == 5:  # 第5个推广引流的a跳过
                    continue
                links = driver.find_elements_by_xpath(
                    "//*[@id='beike']/div[1]/div[4]/div[1]/div[4]/ul//li/a")  # 在每次循环内都重新获取a标签，组成列表
                link = links[j]  # 逐一将列表里的a标签赋给link
                url = link.get_attribute('href')  # 提取a标签内的链接，注意这里提取出来的链接是字符串
                print(url)
                driver.get(url)  # url为字符串，直接用浏览器打开网址即可
                time.sleep(1)  # 留出加载时间
                total = driver.find_element_by_xpath(
                    "//*[@id='beike']/div[1]/div[4]/div[1]/div[2]/div[2]/span[1]").text  # .text的意思是指输出这里的纯文本内容
                print("总价" + total + "万")
                unitPrice = driver.find_element_by_xpath(
                    "//*[@id='beike']/div[1]/div[4]/div[1]/div[2]/div[2]/div[1]/div[1]/span").text
                print("均价" + unitPrice + "元/平米")
                # 房屋户型
                houseType = driver.find_element_by_xpath("//*[@id='introduction']/div/div/div[1]/div[2]/ul/li[1]").text
                print(houseType)
                # 挂牌时间
                onTime = driver.find_element_by_xpath("//*[@id = 'introduction']/div/div/div[2]/div[2]/ul/li[1]").text
                print(onTime)
                driver.back()  # 后退，返回原始页面目录页
                time.sleep(1)  # 留出加载时间
    else:
        pass

time.sleep(1)  # 暂停1秒钟
print("--------------开始爬取上海房屋数据----------------")
for i in range(4):
    driver.get("https://sh.ke.com/ershoufang/rs" + sh_list[i])
    time.sleep(3)  # 暂停2秒钟
    chushou = driver.find_element_by_xpath("//*[@id='sem_card']/div/div[2]/div[2]/div[2]/div[2]")
    chengjiao = driver.find_element_by_xpath("//*[@id='sem_card']/div/div[2]/div[2]/div[3]/div[2]")
    daikan = driver.find_element_by_xpath("//*[@id='sem_card']/div/div[2]/div[2]/div[4]/div[2]")
    print(bj_list[i] + "正在出售" + chushou.text + " 近90天成交" + chengjiao.text + " 近30天带看" + daikan.text)
    time.sleep(2)  # 暂停2秒钟

'''
# 第三方 SMTP 服务
# mail_host="smtp.XXX.com"  #设置服务器
mail_host = "mail.lizihang.com"
mail_user = "shenjing@lizihang.com"  # 用户名
mail_pass = "landz@123"  # 口令

sender = 'shenjing@lizihang.com'  # 发送邮件
receivers = ['liruoyun@lizihang.com', 'shenjing@lizihang.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱

mail_msg = """

<h2>链家网房价</h2>
<p>""" + houseText + """</p>
<p>""" + custText + """</p>
<p>""" + daikanTest + """</p>

<p><a href="https://bj.lianjia.com/fangjia/">这是链家网房价链接</a></p>"""
# message = MIMEText('Python 邮件发送测试...', 'plain', 'utf-8')
message = MIMEText(mail_msg, 'html', 'utf-8')
message['From'] = Header("邮件自动发送--链家网房价", 'utf-8')
message['To'] = Header("李若赟", 'utf-8')

subject = '链家网房价'
message['Subject'] = Header(subject, 'utf-8')

try:
    smtpObj = smtplib.SMTP()
    smtpObj.connect(mail_host, 25)  # 25 为 SMTP 端口号
    smtpObj.login(mail_user, mail_pass)
    smtpObj.sendmail(sender, receivers, message.as_string())
    print("邮件发送成功")
except smtplib.SMTPException:
    print("Error: 无法发送邮件")
'''
driver.quit()
end = time.clock()  # 结束时间
print("运行耗时", end - start)

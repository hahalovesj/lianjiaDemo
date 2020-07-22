#!/usr/bin/python3
 
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from selenium import webdriver # 导入webdriver包
import time
 
#driver = webdriver.Firefox() # 初始化一个火狐浏览器实例：driver
driver = webdriver.Chrome()

driver.maximize_window() # 最大化浏览器 

time.sleep(1) # 暂停5秒钟

driver.get("https://bj.lianjia.com/fangjia/") # 通过get()方法，打开一个url站点

#找到在售房源
#driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[1]/div[1]/div[2]/div[2]/span[4]/a[1]").
ershoufang = driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[1]/div[1]/div[2]/div[2]/span[4]/a[1]")
print(ershoufang.text)
#找到新增房
xinzengfang1 = driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[1]/div[2]/div[1]/div[1]")
xinzengfang2 = driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[1]/div[2]/div[1]/div[2]")
houseText = xinzengfang1.text + xinzengfang2.text
print(houseText)
#找到新增客
xinzengke1 = driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[1]/div[2]/div[2]/div[1]")
xinzengke2 = driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[1]/div[2]/div[2]/div[2]")
custText = xinzengke1.text + xinzengke2.text
print(custText)

#找到带看量
daikan1 = driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[1]/div[2]/div[3]/div[1]")
daikan2 = driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[1]/div[2]/div[3]/div[2]")
daikanTest = daikan1.text + daikan2.text
print(daikanTest)


# 第三方 SMTP 服务
# mail_host="smtp.XXX.com"  #设置服务器
mail_host="mail.lizihang.com"
mail_user="shenjing@lizihang.com"    #用户名
mail_pass="landz@123"   #口令 
 
 
sender = 'shenjing@lizihang.com'#发送邮件
receivers = ['liruoyun@lizihang.com','shenjing@lizihang.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
 
mail_msg = """

<h2>链家网房价</h2>
<p>"""+houseText+"""</p>
<p>"""+custText+"""</p>
<p>"""+daikanTest+"""</p>

<p><a href="https://bj.lianjia.com/fangjia/">这是链家网房价链接</a></p>"""
#message = MIMEText('Python 邮件发送测试...', 'plain', 'utf-8')
message = MIMEText(mail_msg, 'html', 'utf-8')
message['From'] = Header("邮件自动发送--链家网房价", 'utf-8')
message['To'] =  Header("李若赟", 'utf-8')
 
subject = '链家网房价'
message['Subject'] = Header(subject, 'utf-8')
 
try:
    smtpObj = smtplib.SMTP() 
    smtpObj.connect(mail_host, 25)    # 25 为 SMTP 端口号
    smtpObj.login(mail_user,mail_pass)
    smtpObj.sendmail(sender, receivers, message.as_string())
    print ("邮件发送成功")
except smtplib.SMTPException:
    print ("Error: 无法发送邮件")

driver.quit()
 
 
 

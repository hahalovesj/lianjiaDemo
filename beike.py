#coding=utf-8
from appium import webdriver
import time
def get_driver():
    capabilities = {
        "app": "/Users/shine/Downloads/powerapp_202007101205.apk",
        "platformName": "Android",
        "deviceName": "CFM7N18A18000471",
        "appActivity": "com.lizihang.powerapp.activities.SplashActivity",
        "noReset": "true"
    }
    driver = webdriver.Remote("http://localhost:4723/wd/hub",capabilities)
    time.sleep(10)
    return driver

def get_size():
    size = driver.get_window_size()
    width = size['width']
    height = size['height']
    return width,height

#向左边滑动   
def swipe_left():
    x1 = get_size()[0]/10*9
    y1 = get_size()[1]/2
    x = get_size()[0]/10*1
    driver.swipe(x1,y1,x,y1)

#向上边滑动，靠定位元素高度计算，有一点点片差   
def swipe_up():
    x1 = get_size()[0]/2
    y1 = 928
    y2 = 437.5
    driver.swipe(x1,y1,x1,y2,5000)
    time.sleep(2)

#向上边滑动,靠手工测算的，不行  
def swipe_up2():
    x1 = get_size()[0]/2
    y1 = 820
    y2 = 420
    driver.swipe(x1,y1,x1,y2,5000)
    time.sleep(2)    


#滚动,scroll() 与swipe()的区别，swipe是可以根据自己需要设置滑动的距离，而scroll是根据页面中两个元素位置距离进行滑动。
def scroll_page(list):
	#获取倒数第一个和第一个元素，将页面向上滑动
    stop_element=list[1]
    start_element=list[2]
    driver.scroll(start_element,stop_element,3000) 
    time.sleep(5)

#是否滑动到页面底部
def isPageEnd(list):
    # 获取滑动前页面元素
    before_swipe = driver.page_source
    # 滑动
    swipe_up()
    # 获取滑动后页面元素
    after_swipe = driver.page_source
    # 对比滑动前后的页面元素
    # 若滑动前后页面元素完全相同，则滑动失败，已达页尾
    if before_swipe == after_swipe:
        return False
    # 若滑动前后页面元素不同，则滑动成功，继续滑动
    else:
        return True

#TODO https://www.jianshu.com/p/9b6d851a766a
def scrollToElement(): #使用scrollIntoView方法实现滚动到指定控件元素
    webElement = driver.find_element_by_android_uiautomator(
    "new UiScrollable(new UiSelector().scrollable(true)).scrollIntoView(new UiSelector().text(\"Display\"))")
        
       

driver = get_driver()    
#进入新房
driver.find_element_by_id("com.lizihang.powerapp:id/tv_onehand").click()
time.sleep(5)
#获取新房列表
list = driver.find_elements_by_id("com.lizihang.powerapp:id/oneLl")
#time.sleep(5)
#print(list[0].get_attribute("text"))
#进入第一个新房详情
#list[0].click()
#time.sleep(5)
#楼盘名称
#resblockname = driver.find_element_by_id("com.lizihang.powerapp:id/resblockName").get_attribute("text")
#楼盘价格
#totalPriceRange = driver.find_element_by_id("com.lizihang.powerapp:id/totalPriceRange").text
#print(resblockname)
#print(totalPriceRange)
#爬取数据保存到文件，a+附加读写方式打开，若文件不存在，创建。
#fileOb = open('/Volumes/SD/workSpace/vsCodeWorkplace/appiumStudy/powerapp.txt','a+',encoding='utf-8')     
##fileOb.write(resblockname+totalPriceRange)
#time.sleep(5)
#退出详情
#driver.find_element_by_id("com.lizihang.powerapp:id/iv_title_bar_left_icon").click()
#time.sleep(5)
#新房列表滑动
for i in range(10):
    #scroll_page(list)
    swipe_up()
    print ('第'+str(i)+'次滑动~')
#time.sleep(5)


driver.quit()

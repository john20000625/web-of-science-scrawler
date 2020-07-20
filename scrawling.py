from selenium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import re
import os
from selenium.common.exceptions import NoSuchElementException
numberlist=[]
flag=True
def txt_xls(filename,xlsname):
    try:
        f=open(filename)
        xls=xlwt.Workbook()
        #生成excel的方法，声明excel
        sheet=xls.add_sheet('sheet1',cell_overwrite_ok=True)
        x=0
        while True:
            #按行循环，读取文本文件
            line=f.readline()
            if not line:
                break #如果没有内容，则退出循环
            for i in range(len(line.split('\t'))):
                item=line.split('\t')[i]
                sheet.write(x,i,item) #x单元格精度，i单元格纬度
            x+=1 #excel另起一行
        f.close()
        xls.save(xlsname)#保存xls文件
    except:
        raise
def click_by_time(driver,xpath,maxTime):#防止动态元素未加载出来，无法点击导致程序出错，故定义该方法，等元素加载出来再点击它
    t = 0
    while t<=maxTime:
        if driver.find_element_by_xpath(xpath)!=None: 
            driver.find_element_by_xpath(xpath).click()
            break
        time.sleep(1)
        t+=1  
filepath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))+'/namelist.txt'
n=open(filepath,'r')
Lines=n.readlines()
for i in range(len(Lines)):
    if flag==False:
        for j in range(len(Lines)):
            if '!' not in Lines[j]:
                i=j+1
                break
    keyword = Lines[i].strip('\n')    #搜索的关键词，到时候for循环就不停更改这个
    driver = webdriver.Chrome()
    driver.get('https://apps.webofknowledge.com/')

    #先搞定WOS，把数据库选择成wos核心数据库，这样url里才会带wos
    click_by_time(driver,'//*[@id="select.database.stripe"]/div/div/span[2]/span[1]/span/span[2]',10)
    click_by_time(driver,'/html/body/span[27]/span/span[2]/ul/li[2]',10)
    time.sleep(5)
    #找到搜索框
    searchingBox = driver.find_element_by_xpath('//*[@id="value(input1)"]')
    searchingBox.send_keys(keyword)

    #找到年份选择，并点击自定义年份（使用坐标）
    yearBox = driver.find_element_by_xpath('//*[@id="timespan"]/div[2]/div/span/span[1]/span/span[2]/b')
    yearBox.click()
    time.sleep(2)
    ActionChains(driver).move_by_offset(124,420).click().perform()
    time.sleep(3)
    #左边年份，从xxx年开始
    leftYear = driver.find_element_by_xpath('//*[@id="timespan"]/div[3]/div/span[2]/span[1]/span')
    leftYear.click()
    leftBox = driver.find_element_by_xpath('/html/body/span[37]/span/span[1]/input')
    ActionChains(driver).move_to_element(leftBox).send_keys('2019').send_keys(Keys.ENTER).perform()

    #右边年份，到xxx年结束
    rightYear = driver.find_element_by_xpath('//*[@id="timespan"]/div[3]/div/span[4]/span[1]/span')
    rightYear.click()
    rightBox = driver.find_element_by_xpath('/html/body/span[37]/span/span[1]/input')
    ActionChains(driver).move_to_element(rightBox).send_keys('2019').send_keys(Keys.ENTER).perform()

    #点选按照出版物名称搜索
    theme = driver.find_element_by_xpath('//*[@id="searchrow1"]/td[2]/span/span[1]/span')
    theme.click()
    click_by_time(driver,'//*[@id="select2-select1-results"]/li[4]',10) #尤其注意，这里是动态id,不能直接复制按照id抓，所以要自己修改xpath

    #点击搜索
    searchingButton = driver.find_element_by_xpath('//*[@id="searchCell1"]/span[1]/button')
    searchingButton.click()

    #获取当前url，后面其实用不上
    currentPageUrl = driver.current_url
    #print(currentPageUrl)



    #--------divide---------

    #初始化并打开网页
    #driver2 = webdriver.Chrome()
    #driver2.get(currentPageUrl)
    

    #先确定杂志是否存在，是的话点击导出，并选择其他保存方式，否则关闭开启下一轮循环
    try:
        click_by_time(driver,'//*[@id="DocumentType_img"]',10)
        click_by_time(driver,'//*[@id="DocumentType_1"]',10)
        click_by_time(driver,'//*[@id="DocumentType_tr"]/button[1]',10)
    except NoSuchElementException:
        flag=False
        driver.quit()
        continue
    
    click_by_time(driver,'//*[@id="exportTypeName"]',10)
    click_by_time(driver,'//*[@id="saveToMenu"]/li[3]/a',10)
    #爬取检索结果数量
    number = driver.find_element_by_xpath('//*[@id="page"]/div[1]/div[26]/div[2]/div/div/div/div[1]/div[1]/div/div/h3/span')
    txt_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + 'numberlist2019.txt'
    f = open(txt_path,'a+')
    f.write(number.text)
    f.write('\n')
    f.close()
    #点击难处理的Records from按钮
    #time.sleep(2)#让网页加载一会儿
    #button1 = driver.find_element_by_xpath('//*[@id="numberOfRecordsRange"]') #找到Records from的location
    #location = button1.location
    #print(location)
    #ActionChains(driver).move_by_offset(550,255).click().perform() #在找到的坐标上要偏移十几个像素点才行，作废，没必要用坐标，麻烦
    click_by_time(driver,'//*[@id="numberOfRecordsRange"]',10)

    #点击选择全引文下载
    click_by_time(driver,'//*[@id="select2-bib_fields-container"]',10)
    click_by_time(driver,'//*[@id="select2-bib_fields-results"]/li[4]',10)

    #点击下载
    click_by_time(driver,'//*[@id="exportButton"]',10)
    time.sleep(10)
    driver.quit()
filename1 = txt_path
xlsname1 = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + 'numberlist2019.xls'
txt_xls(filename1,xlsname1)
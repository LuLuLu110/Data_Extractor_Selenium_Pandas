from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import glob
import time
import os
import pandas as pd




def download_web_data():


    # 确定chromedriver的路径
    driver_path = os.path.abspath("chromedriver.exe")

    # 创建一个新的web driver对象，并为他命名为browser
    browser = webdriver.Chrome(driver_path)
    # 给browser设置初始访问的url
    url = "https://github.com/Frankdew1995/Python_Cookbook_For_Lu"
    browser.get(url)
    # 最大化浏览器窗口
    browser.maximize_window()


    # 在网页找到操作对象的Xpath，并点击
    WebDriverWait(browser, 45).until(EC.presence_of_element_located((By.XPATH,
                                                                     '//*[@id="c92e9ae3a846ea015129b8dbda266dec-a04690aa1d4a8a262676be4c9d7d1b82487b4b81"]'))).click()

    WebDriverWait(browser, 45).until(EC.presence_of_element_located((By.XPATH,
                                                                     '//*[@id="b9b05beafa24a07f43b95b28df16a350-5bcf2fd2047135411e942810a05243f0ddb3666e"]'))).click()


    # 点击下载
    WebDriverWait(browser, 45).until(EC.presence_of_element_located((By.XPATH,
                                                                     '//*[@id="raw-url"]'))).click()
    time.sleep(15)
    # 下载完毕关闭浏览器
    browser.close()








def DataPrices():
    # 切换路径
    os.chdir("/Users/frankdu/Downloads")

    print(os.getcwd())
    # 使用glob获取所以excel文件
    files = glob.glob("*.xlsx")
    # 排序查找最新下载文件，并读取
    files.sort(key=os.path.getmtime)
    newFile = files[-1]
    df=pd.read_excel(newFile,sheet_name=0)

    # 切换到初始路径
    os.chdir("/Users/frankdu/Data_Extractor")
    # 读取底表数据
    df1 = pd.read_excel("Products.xlsx",sheet_name=0)

    # 根据df1的"EAN"与df合并
    df2 = pd.merge(df1, df, on="EAN", how="left")

    # 将合并以后的df2生成excel文件
    df2.to_excel("UpdatedPrice.xlsx", index=False)


# 运行两个定义的函数
#运行Selenium函数自动下载数据
download_web_data()
#使用Pandas本地处理并合并数据
DataPrices()



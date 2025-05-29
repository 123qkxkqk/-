import tkinter as tk
from tkinter import messagebox
import math
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service  # 新增导入
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def start(request_num):
    start_time = time.time()  # 开始计时
    data = pd.read_excel('a1.xlsx', sheet_name='Sheet1')
    # 指定ChromeDriver路径
    chrome_driver_path = r"D:\pycharm\PyCharm Community Edition 2025.1\bin\chromedriver.exe"  # 修改为你的实际路径
    service = Service(executable_path=chrome_driver_path)  # 新增Service对象
    # 配置调试模式
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print("----成功连接到浏览器----")
    print("当前页面标题:", driver.title)
    print(f"当前窗口句柄: {driver.current_window_handle}")
    print(f"当前窗口地址：{driver.current_url}")
    # 获取所有窗口句柄
    all_handles = driver.window_handles
    print(f"当前共有 {len(all_handles)} 个窗口")

    for i in range(21):
        if i in {5, 19, 18, 20, 6, 7, 16, 17}:  # 竖版大图需要传logo
            for k, handle in enumerate(all_handles):
                driver.switch_to.window(handle)
                option = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[.//span[contains(text(), "本页全部添加")]]')))
                driver.execute_script("arguments[0].click();", option)  # 点击
                time.sleep(0.5)
                try:
                    # 删除多余图片
                    element = driver.find_element(By.XPATH,
                                                  '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[2]/div/div[2]/div/div[3]/table/tbody/tr[1]/td[1]/div')
                    a = int(element.text)  # 提取到的内容转化为数值型
                    # 去除多余的素材
                    for j in range(a - request_num):
                        option = driver.find_element(By.XPATH,
                                                     '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[2]/div/div[2]/div/div[3]/table/tbody/tr[1]/td[2]/div/div/div/div[2]/span')
                        driver.execute_script("arguments[0].click();", option)  # 点击x号
                    # 时间缓和
                    time.sleep(0.5)
                    # 生成品牌logo
                    driver.find_element(By.XPATH,
                                    '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[1]/div/div[2]/div/div[1]/form/div[1]/div[1]/div/div/div/div/span/span/i').click()
                    if i in {5, 19}:
                        option = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "品牌logo")]')))
                        driver.execute_script("arguments[0].click();", option)
                    elif i in {18, 20}:
                        option = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "视频封面图")]')))
                        driver.execute_script("arguments[0].click();", option)
                    elif i in {6, 7, 16, 17}:
                        option = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "视频结束页背景")]')))
                        driver.execute_script("arguments[0].click();", option)

                    if k+1 == len(all_handles):
                        time.sleep(3)
                    else:
                        time.sleep(1)
                except Exception as e:
                    # 时间缓和
                    time.sleep(2)
                    # 生成品牌logo
                    driver.find_element(By.XPATH,
                                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[1]/div/div[2]/div/div[1]/form/div[1]/div[1]/div/div/div/div/span/span/i').click()
                    if i in {5, 19}:
                        option = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "品牌logo")]')))
                        driver.execute_script("arguments[0].click();", option)
                    elif i in {18, 20}:
                        option = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "视频封面图")]')))
                        driver.execute_script("arguments[0].click();", option)
                    elif i in {6, 7, 16, 17}:
                        option = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "视频结束页背景")]')))
                        driver.execute_script("arguments[0].click();", option)

                    if k+1 == len(all_handles):
                        time.sleep(4)
                    else:
                        time.sleep(0.5)
                    continue

            # 重新迭代窗口 添加logo
            for k, handle in enumerate(all_handles):
                driver.switch_to.window(handle)
                # 点击"全部添加"
                option = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[.//span[contains(text(), "本页全部添加")]]')))
                driver.execute_script("arguments[0].click();", option)  # 点击
                time.sleep(0.5)
                try:
                    element = driver.find_element(By.XPATH,
                                                  '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[2]/div/div[2]/div/div[3]/table/tbody/tr[1]/td[1]/div')
                    a = int(element.text)  # 提取到的内容转化为数值型
                    # 去除多余的素材
                    for n in range(a - request_num):
                        driver.find_element(By.XPATH,
                                                     '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[2]/div/div[2]/div/div[3]/table/tbody/tr[1]/td[3]/div/div/div/div[2]/span').click()
                    time.sleep(0.5)
                    #点击下一个版位
                    option = driver.find_element(By.XPATH,
                                                 f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[{i + 2}]')
                    # 使用JavaScript点击//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[1]
                                           #//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[21]
                    driver.execute_script("arguments[0].click();", option)
                    time.sleep(0.5)
                except Exception as e:
                    time.sleep(2)
                    # 点击下一个版位
                    option = driver.find_element(By.XPATH,
                                                 f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[{i + 2}]')
                    # 使用JavaScript点击
                    driver.execute_script("arguments[0].click();", option)
                    time.sleep(0.5)
                    continue

        else:
            # 循环切换窗口
            for k, handle in enumerate(all_handles):
                driver.switch_to.window(handle)
                option = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[.//span[contains(text(), "本页全部添加")]]')))
                driver.execute_script("arguments[0].click();", option)  # 点击"本页全部添加"
                time.sleep(1)
                # 删除多余图片
                try:
                    element = driver.find_element(By.XPATH,
                                                  '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[2]/div/div[2]/div/div[3]/table/tbody/tr[1]/td[1]/div')
                    a = int(element.text)  # 提取到的内容转化为数值型
                    # 去除多余的素材
                    for j in range(a - request_num):
                        option = driver.find_element(By.XPATH,
                                                     '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[2]/div/div[2]/div/div[3]/table/tbody/tr[1]/td[2]/div/div/div/div[2]/span')
                        driver.execute_script("arguments[0].click();", option)
                    # 点击下一个版位
                    time.sleep(1)
                    if i < 21:  # 最后一个版位不用点击
                        option = driver.find_element(By.XPATH,
                                                     f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[{i + 2}]')
                        # 使用JavaScript点击
                        driver.execute_script("arguments[0].click();", option)
                    # 窗口切换时间缓和
                    if k + 1 == len(all_handles):
                        time.sleep(4)
                    else:
                        time.sleep(0.5)
                except Exception as e:
                    print(f"窗口 {k + 1} 操作异常：{str(e)}")
                    # 点击下一个版位
                    time.sleep(2)
                    option = driver.find_element(By.XPATH,
                                                 f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[{i + 2}]')
                    # 使用JavaScript点击
                    driver.execute_script("arguments[0].click();", option)
                    time.sleep(2)
                    continue

    # 在程序结束位置计算运行时间
    end_time = time.time()
    total_seconds = end_time - start_time
    # 转换成分秒格式
    minutes = math.floor(total_seconds / 60)
    seconds = int(total_seconds % 60)
    print(f"{len(all_handles)}条广告计划上传创意完成时间：{minutes}分{seconds}秒")
    return total_seconds  # 新增返回值

if __name__ == "__main__":
    start(5)
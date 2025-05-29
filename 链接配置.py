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
start_time = time.time()  #开始计时
###初始化素材数量
request_num = 10
#初始化任务索引：
index = 42
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

for k, handle in enumerate(all_handles):
    driver.switch_to.window(handle)
    # 获取窗口信息
    print(f"【第 {k + 1} 个窗口】")
    m = k + index #设置一个分任务索引
    # 输入投放产品名称后等待选项出现
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[2]/div/form/div[2]/div/div[1]/div[1]/input').send_keys(
        data.iloc[m]['产品名称'])
    # 显式等待下拉选项加载并且输入产品名称（最多等待10秒）
    option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{str(int(data.iloc[m]['产品ID']))}')]")))
    option.click()
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[2]/div/form/div[1]/div/div/label[1]/span[1]/span').click()
    # 添加快游戏链接
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[2]/div/form/div[3]/div/div/div[1]/div/form/div[1]/div/div/form/div/div[2]/div[1]/div/div[1]/input').send_keys(
        data.iloc[m]['快应用链接'])
    # 更改为落地页
    driver.find_element(By.XPATH, '//input[@placeholder="请选择"]').click()
    print("✅ 已展开链接类型下拉框")

    ##方案一：
    # option = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "自定义落地页")]')))
    # option.click()

    option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//li[contains(.//text(), "自定义落地页")]')))
    option.click()

    # 添加落地页
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[2]/div/form/div[3]/div/div/div[1]/div/form/div[1]/div/div/form/div/div[2]/div[2]/div/div[2]/div/div/input').send_keys(
        data.iloc[m]['落地页链接'])
    # 点击"已验证"按钮
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[2]/div/form/div[3]/div/div/div[1]/div/form/div[2]/div/button/span').click()
    # 方案一：时间休息硬着陆
    # time.sleep(3)
    # 方案二：显性等待
    # option = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//button[contains(.//text(), "已验证")]')))
    option = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//button[.//span[contains(text(), "已验证")]]')))
    # 点击下一步按钮
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[1]/button[3]/span').click()
    # 创意页面加载
    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.CLASS_NAME, "el-loading-mask")))
    print("✅ 页面加载完成")
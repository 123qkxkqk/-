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
request_num = 3
#初始化任务索引：
index = 45
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
    m = k + index  # 设置一个分任务索引
    # 点击批量创建
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div/div[2]/div[1]/div[1]/button/span').click()
    # 输入模板账户后等待选项出现
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/form/div[1]/div/div/div/input').send_keys(
        data.iloc[m]['模板账户'])
    # 显式等待下拉选项加载（最多等待10秒）
    option = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//li[contains(text(), '{data.iloc[m]['模板账户']}')]")))
    option.click()
    time.sleep(1)
    # 点击快应用
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/form/div[2]/div/div/label[2]/span[1]/span').click()
    time.sleep(2)
    # 输入模板计划后等待选项出现
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/form/div[3]/div/div/div/div[2]/input').send_keys(
        data.iloc[m]['模板计划'])
    # 显式等待下拉选项加载（最多等待10秒）
    time.sleep(1)
    option = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//li[contains(text(), '{data.iloc[m]['模板计划']}')]")))
    option.click()
    # 去除下拉列表
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/form/div[2]/div/div/label[2]/span[1]/span').click()
    # 点击核实计划按钮
    verify_button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '//button[.//span[contains(text(), "核实计划")]]')))
    verify_button.click()

    # 弹窗标题
    title_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located(
            (By.XPATH, '//div[contains(@class, "el-dialog") and contains(.//text(), "核实计划信息")]')))
    print("已经出现弹窗")
    # 点击"已核实"按钮
    confirmed_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//div[@class="el-dialog__body"]//button[.//span[contains(text(), "已核实")]]')))
    confirmed_button.click()
    print("✅ 已确认核实")
    # 删除已有的目标账户
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/form/div[5]/div/div[1]/div[1]/span/span/i').click()
    # 创建新的目标账户
    driver.find_element(By.XPATH,
                        '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/form/div[5]/div/div[1]/div[1]/input').send_keys(
        str((data.iloc[m]['目标账户'])))
    option = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//li[contains(text(), '{str((data.iloc[m]['目标账户']))}')]")))
    option.click()

    # 点击"下一步"按钮
    driver.find_element(By.XPATH,
                     '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[1]/button[2]/span').click()

# 在程序结束位置计算运行时间
end_time = time.time()
total_seconds = end_time - start_time
# 转换成分秒格式
minutes = math.floor(total_seconds / 60)
seconds = int(total_seconds % 60)
print(f"{len(all_handles)}条广告计划创建完成时间：{minutes}分{seconds}秒")
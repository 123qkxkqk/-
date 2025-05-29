import time

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service  # 新增导入
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
    time.sleep(3)
    driver.switch_to.window(handle)
    print(f"窗口{k+1}:{driver.title}")
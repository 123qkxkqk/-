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

#配置目标账户
def create_id(request_num, index):
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

        # 点击"已核实"按钮
        confirmed_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//div[@class="el-dialog__body"]//button[.//span[contains(text(), "已核实")]]')))
        confirmed_button.click()

        # 删除已有的目标账户
        driver.find_element(By.XPATH,
                            '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/form/div[5]/div/div[1]/div[1]/span/span/i').click()
        # 创建新的目标账户
        driver.find_element(By.XPATH,
                            '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/form/div[5]/div/div[1]/div[1]/input').send_keys(
            str((data.iloc[m]['目标账户'])))
        option = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//li[contains(text(), '{str((data.iloc[m]['目标账户'])).strip()}')]")))
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
    print(f"{len(all_handles)}条目标账户配置完成时间：{minutes}分{seconds}秒")
    return total_seconds  # 新增返回值

#配置链接：
def add_url(request_num, index):
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

    for k, handle in enumerate(all_handles):
        m = k + index
        driver.switch_to.window(handle)
        # 获取窗口信息
        print(f"【第 {k + 1} 个窗口】")
        m = k + index  # 设置一个分任务索引
        # 输入投放产品名称后等待选项出现
        driver.find_element(By.XPATH,
                            '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[2]/div/form/div[2]/div/div[1]/div[1]/input').send_keys(
            data.iloc[m]['产品名称'].strip())
        # 显式等待下拉选项加载并且输入产品名称（最多等待10秒）
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{str(data.iloc[m]['产品ID'])}')]")))
        option.click()
        driver.find_element(By.XPATH,
                            '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[2]/div/form/div[1]/div/div/label[1]/span[1]/span').click()
        # 添加快游戏链接
        driver.find_element(By.XPATH,
                            '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/div[2]/div/form/div[3]/div/div/div[1]/div/form/div[1]/div/div/form/div/div[2]/div[1]/div/div[1]/input').send_keys(
            data.iloc[m]['快应用链接'])
        # 更改为落地页
        driver.find_element(By.XPATH, '//input[@placeholder="请选择"]').click()

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
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//button[.//span[contains(text(), "已验证")]]')))
        # 点击下一步按钮
        driver.find_element(By.XPATH,
                            '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[1]/button[3]/span').click()
        # 创意页面加载
        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.CLASS_NAME, "el-loading-mask")))

    # 在程序结束位置计算运行时间
    end_time = time.time()
    total_seconds = end_time - start_time

    # 转换成分秒格式
    minutes = math.floor(total_seconds / 60)
    seconds = int(total_seconds % 60)
    print(f"{len(all_handles)}条广告计划配置连接时间：{minutes}分{seconds}秒")
    return total_seconds  # 新增返回值
#上传创意
def add_ideals(request_num, index):
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

    for i in range(9):
        if i == 3:
            pass
        elif i == 6:  # 竖版大图需要传logo
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
                    option = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "品牌logo")]')))
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
                    option = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//li[contains(.//text(), "品牌logo")]')))
                    driver.execute_script("arguments[0].click();", option)
                    if k+1 == len(all_handles):
                        time.sleep(4)
                    else:
                        time.sleep(0.5)
                    continue

            # 重新迭代窗口 添加logo
            for k, handle in enumerate(all_handles):
                m = k + index
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
                        option = driver.find_element(By.XPATH,
                                                     '//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[4]/div/div[2]/div/div[2]/div/div[3]/table/tbody/tr[1]/td[3]/div/div/div/div[2]/span').click()
                    time.sleep(0.5)
                    #点击下一个版位
                    if data.iloc[m]['模板计划'] == "长喵了-ROI1-不限版位-WK":
                        option = driver.find_element(By.XPATH,
                                                     f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[{i + 2}]')
                        # 使用JavaScript点击
                        driver.execute_script("arguments[0].click();", option)
                    elif data.iloc[m]['模板计划'] == "长喵了-Costcap-ROI1-WK":
                        option = driver.find_element(By.XPATH,
                                                     f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[4]/div/div/div[{i + 2}]')
                        # 使用JavaScript点击
                        driver.execute_script("arguments[0].click();", option)
                    time.sleep(0.5)
                except Exception as e:
                    time.sleep(2)
                    # 点击下一个版位
                    if data.iloc[m]['模板计划'] == "长喵了-ROI1-不限版位-WK":
                        option = driver.find_element(By.XPATH,
                                                     f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[{i + 2}]')
                        # 使用JavaScript点击
                        driver.execute_script("arguments[0].click();", option)
                    elif data.iloc[m]['模板计划'] == "长喵了-Costcap-ROI1-WK":
                        option = driver.find_element(By.XPATH,
                                                     f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[4]/div/div/div[{i + 2}]')
                        # 使用JavaScript点击
                        driver.execute_script("arguments[0].click();", option)
                    time.sleep(0.5)
                    continue
        else:
            # 循环切换窗口
            for k, handle in enumerate(all_handles):
                m = k + index
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
                    if i < 8:  # 最后一个版位不用点击
                        if data.iloc[m]['模板计划'] == "长喵了-ROI1-不限版位-WK":
                            option = driver.find_element(By.XPATH,
                                                         f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[{i + 2}]')
                            # 使用JavaScript点击
                            driver.execute_script("arguments[0].click();", option)
                        elif data.iloc[m]['模板计划'] == "长喵了-Costcap-ROI1-WK":
                            option = driver.find_element(By.XPATH,
                                                         f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[4]/div/div/div[{i + 2}]')
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
                    if i < 8:  # 最后一个版位不用点击
                        if data.iloc[m]['模板计划'] == "长喵了-ROI1-不限版位-WK":
                            option = driver.find_element(By.XPATH,
                                                         f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[3]/div/div/div[{i + 2}]')
                            # 使用JavaScript点击
                            driver.execute_script("arguments[0].click();", option)
                        elif data.iloc[m]['模板计划'] == "长喵了-Costcap-ROI1-WK":
                            option = driver.find_element(By.XPATH,
                                                         f'//*[@id="app"]/div[1]/div[2]/section/div/div/div[2]/div/div/div[3]/div[2]/div/form/div/div[2]/form/div[4]/div/div/div[{i + 2}]')
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

#弹出窗口函数
def show_confirmation_dialog(title, message):
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.attributes('-topmost', True)  # 确保弹窗在最前面
    messagebox.showinfo(title, message)
    root.destroy()

#时间格式化函数
def format_time(t):
    minute = math.floor(t / 60)
    second = int(t % 60)
    return f"{minute}分{second}秒"

#交互弹窗函数
def get_input_values():
    """创建输入弹窗获取a和b的值"""
    root = tk.Tk()
    root.title("参数输入")
    root.geometry("300x150")
    # 设置输入标签和输入框
    tk.Label(root, text="请输入素材数量(a):").grid(row=0, padx=10, pady=5)
    entry_a = tk.Entry(root)
    entry_a.grid(row=0, column=1, padx=10, pady=5)
    tk.Label(root, text="请输入起始行号(b):").grid(row=1, padx=10, pady=5)
    entry_b = tk.Entry(root)
    entry_b.grid(row=1, column=1, padx=10, pady=5)
    result = []  # 用于存储输入结果

    def on_confirm():
        """验证并获取输入值"""
        try:
            a_val = int(entry_a.get())
            b_val = int(entry_b.get())
            if a_val <= 0 or b_val <= 0:
                raise ValueError
            result.extend([a_val, b_val])
            root.destroy()
        except ValueError:
            tk.messagebox.showerror("输入错误", "请输入有效的正整数")
            entry_a.delete(0, tk.END)
            entry_b.delete(0, tk.END)
    # 确认按钮
    btn_confirm = tk.Button(root, text="确认", command=on_confirm)
    btn_confirm.grid(row=2, columnspan=2, pady=10)
    # 使窗口居中
    root.eval('tk::PlaceWindow . center')
    root.mainloop()
    return result[0], result[1] if result else (None, None)

def start(request_num, index):
    t1 = create_id(request_num, index)
    show_confirmation_dialog("确认", "配置账户完成，请观察是否有其他弹窗")
    t2 = add_url(request_num, index)
    show_confirmation_dialog("确认", "链接配置完成，请筛选素材渠道或者选择产品名称")
    t3 = add_ideals(request_num, index)
    #计算总耗时
    total_time = t1 + t2 + t3
    show_confirmation_dialog(
        "确认",
        f"所有操作已完成！ 总耗时：{format_time(total_time)}\n"
        f"各阶段耗时：\n"
        f"- 账户配置：{format_time(t1)}\n"
        f"- 链接配置：{format_time(t2)}\n"
        f"- 创意上传：{format_time(t3)}"
    )
if __name__ == "__main__":
    # 获取用户输入
    a, b = get_input_values()
    # 验证输入有效性
    if a and b:
        start(a, b-1)
    else:
        print("输入取消或无效，程序终止")
        show_confirmation_dialog("错误", "输入无效，程序已终止")

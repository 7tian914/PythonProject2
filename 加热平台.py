# ====================================================================================================================
# 数据{各种余额：balance_dict  优惠券：total  订单：record  视频号名称：combined_name  微信豆流水：wechatcoin}
# ================================================================================================================================
import time
import os
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import pandas as pd
import glob
from tabulate import tabulate
import uuid
import shutil

# 配置参数
download_dir = r"C:\Users\benxing\Downloads"
os.makedirs(download_dir, exist_ok=True)
# 创建唯一的下载目录
session_id = str(uuid.uuid4())[:8]  # 生成唯一标识符
temp_download_dir = os.path.join(download_dir, f"temp_{session_id}")
os.makedirs(temp_download_dir, exist_ok=True)
# 浏览器设置
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
chrome_options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
# 修改浏览器配置，使用唯一的下载目录
prefs = {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)
# 初始化浏览器
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

# 全局变量
global name_text,video_account, tongyong, zhuanyong, chongzeng, zhizeng, cps_ad, cps_buy, jiare, total_coupon, remaining_balance
name_text = video_account = ""
tongyong = zhuanyong = chongzeng = zhizeng = cps_ad = cps_buy = jiare = total_coupon = 0
remaining_balance = []


# 保存cookie
def save_cookies(driver, file_path):
    """保存当前Cookies到文件"""
    cookies = driver.get_cookies()

    with open(file_path, 'w') as file:
        json.dump(cookies, file)
    print(f"Cookies已保存到: {file_path}")


# 加载cookie
def load_cookies(driver, file_path):
    """从文件加载Cookies到当前会话"""
    if not os.path.exists(file_path):
        print(f"Cookie文件不存在: {file_path}")
        return False
    with open(file_path, 'r', encoding='utf-8') as file:
        cookies = json.load(file)
    monitor()  # 启动监听
    driver.get("https://channels.weixin.qq.com/promote/pages/platform/login")
    time.sleep(1)
    # 删除旧Cookies
    driver.delete_all_cookies()
    # 添加新Cookies
    for cookie in cookies:
        # 删除domain字段避免兼容性问题
        if 'domain' in cookie:
            del cookie['domain']
        # 处理可能过期的字段
        if 'expiry' in cookie:
            del cookie['expiry']
        try:
            driver.add_cookie(cookie)
        except Exception as e:
            print(f"添加cookie {cookie['name']} 时出错: {e}")

    print(f"已加载 {len(cookies)} 个Cookies")
    return True


# 登录验证
def check_logged_in(driver):
    """检查是否已登录"""
    try:
        # 增加等待时间到15秒，并检查多个可能的登录成功标志
        WebDriverWait(driver, 15).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'新建订单')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(@class,'user-info')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'账户余额')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'订单管理')]"))
            )
        )
        return True
    except TimeoutException:
        print("等待登录超时，未检测到登录成功元素")
        return False
    except Exception as e:
        print(f"检测登录状态时发生异常: {e}")
        return False


def main_automation(driver, download_dir):
    # Cookie文件路径
    cookie_file = os.path.join(download_dir, "wechat_video_cookies.json")
    logged_in = False
    monitor()  # 启动监听

    try:
        # 先尝试加载Cookies免登录
        if os.path.exists(cookie_file):
            print("尝试使用Cookies自动登录...")
            if load_cookies(driver, cookie_file):
                driver.refresh()
                # 增加等待时间
                time.sleep(10)
                if check_logged_in(driver):
                    print("使用Cookies自动登录成功")
                    logged_in = True
                else:
                    print("Cookies已过期，需要重新扫码登录")

        # 如果没有通过Cookies登录成功，走正常登录流程
        if not logged_in:
            print("开始正常登录流程...")
            driver.get("https://channels.weixin.qq.com/promote/pages/platform/login")
            time.sleep(3)
            element = driver.find_element(By.XPATH,
                                          '/html/body/div/div/div/div/div[1]/div[2]/div/div[1]/img')
            screenshot_path = os.path.join(download_dir, "scan.png")
            element.screenshot(screenshot_path)
            print(f"二维码截图已保存至: {screenshot_path}")

            try:
                def wait_for_page_loaded(driver, timeout=30):
                    """组合多种条件检测页面加载完成"""
                    wait = WebDriverWait(driver, timeout)
                    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
                    conditions = [
                        EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'新建订单')]")),
                        EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'最近订单')]"))
                    ]
                    wait.until(EC.any_of(*conditions))

                wait_for_page_loaded(driver)
                print("登录成功，首页加载完成")
            except TimeoutException:
                print("等待超时 - 首页加载")
                return False
        name()  # 获取账户名
        # =================================================余额==================================================================
        money()

        # =================================================优惠券==================================================================
        coupon()

        # =================================================订单管理==================================================================
        time.sleep(3)
        try:
            driver.execute_script(
                'document.querySelectorAll(".finder-ui-desktop-menu__popup_sub_menu")[0].style.display = "flex";')
            element = driver.find_element(By.XPATH,
                                          '/html/body/div/div/div[1]/div[1]/div/div/ul/li[1]/div/ul/li[2]/a/span/span/span')
            driver.execute_script("arguments[0].click();", element)
        except:
            print("找不到")
        time.sleep(5)
        try:
            driver.execute_script('document.querySelector(".finder-ui-desktop-dropdown-menu").style.display = "flex";')
            time.sleep(2)
            driver.find_element(By.XPATH, "//*[@title='加热中']").click()
        except:
            print("失败")
        process_response()  # 获取数据

        # =================================================充值信息==================================================================
        wechat_coins()  # 微信豆明细

        # ==================================================视频号名称获取=============================================================

        return True

    except Exception as e:
        print(f"发生错误: {e}")
        return False
    finally:
        pass


# 各类余额
def money():
    try:
        driver.execute_script(
            'document.querySelectorAll(".finder-ui-desktop-menu__popup_sub_menu")[0].style.display = "flex";')
        element = driver.find_element(By.XPATH,
                                      '/html/body/div/div/div[1]/div[1]/div/div/ul/li[1]/div/ul/li[1]/a/span/span/span')
        driver.execute_script("arguments[0].click();", element)
    except:
        print("找不到")
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, ".user-meta-label")
        balance_dict = {}

        global tongyong, zhuanyong, chongzeng, zhizeng, cps_ad, cps_buy, jiare
        for el in elements:
            text = el.text.strip()
            if text:
                text = text.replace('\n', ' ')
                parts = text.split()
                if len(parts) >= 2:
                    label = ' '.join(parts[:-1])
                    value = parts[-1]
                    balance_dict[label] = value

                else:
                    balance_dict[text] = "0"
        # print(balance_dict)
        # 将余额字典中的值分别赋值给变量
        tongyong = float(balance_dict.get('通用余额', '0'))  # 通用余额
        zhuanyong = float(balance_dict.get('专用余额', '0'))  # 专用余额
        chongzeng = float(balance_dict.get('充赠余额', '0'))  # 充赠余额
        zhizeng = float(balance_dict.get('直赠余额', '0'))  # 直赠余额
        cps_ad = float(balance_dict.get('CPS广告激励余额', '0'))  # CPS广告激励余额
        cps_buy = float(balance_dict.get('CPS内购激励余额', '0'))  # CPS内购激励余额
        jiare = float(balance_dict.get('加热广告激励余额', '0'))  # 加热广告激励余额

        # element = driver.find_element(By.CSS_SELECTOR,
        #                               '.flex.flex-wrap.items-center.gap-x-6')
        # screenshot_path = os.path.join(download_dir, "money.png")
        # element.screenshot(screenshot_path)
        # print(f"二维码截图已保存至: {screenshot_path}")
    except:
        print("未获取到余额")
    time.sleep(3)


# 优惠券
def coupon():
    global total_coupon
    try:
        ol = driver.find_element(By.XPATH,
                                 "/html/body/div/div/div[1]/div[2]/div/div/div[3]/div[1]/div[2]/button")
        driver.execute_script("arguments[0].click();", ol)
    except:
        print('找不到按钮')
    time.sleep(3)
    try:
        al = driver.find_element(By.XPATH,
                                 "(//div[@class='hover-mask'])[1]")
        driver.execute_script("arguments[0].click();", al)
        time.sleep(2)
        bl = driver.find_element(By.XPATH,
                                 "(//button[contains(@type,'button')])[17]")
        driver.execute_script("arguments[0].click();", bl)
        time.sleep(2)
        cl = driver.find_element(By.XPATH,
                                 "/html/body/div/div/div[1]/div[2]/div/div/div[3]/div[1]/form/div[6]/div/div/label[2]")
        driver.execute_script("arguments[0].click();", cl)
        time.sleep(2)
        dl = driver.find_element(By.XPATH,
                                 "/html/body/div/div/div[1]/div[2]/div/div/div[3]/div[1]/form/div[9]/div/div/label[2]/span")
        driver.execute_script("arguments[0].click();", dl)
    except:
        print('找不到按钮')
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, ".value")
        # 提取所有文本并过滤出数字部分
        values = []
        for el in elements:
            text = el.text.strip()
            # 提取数字部分（去掉"点"等非数字字符）
            num = ''.join(filter(str.isdigit, text))
            if num:  # 确保是有效数字
                values.append(int(num))

        total = sum(values)
        total_coupon = float(total)  # 转换为float类型
        print(f"优惠券总点数: {total}点")
        # element = driver.find_element(By.XPATH,
        #                               '//div[contains(@class,"mb-5 mt-[-26px] flex flex-wrap pl-[178px]")]')
        # screenshot_path = os.path.join(download_dir, "coupon.png")
        # element.screenshot(screenshot_path)
        # print(f"二维码截图已保存至: {screenshot_path}")
    except:
        print('找不到优惠券或计算总点数失败')
        print('找不到优惠券')


# 启用网络请求监听
def monitor():
    # 启用网络监听
    driver.execute_cdp_cmd('Network.enable', {})
    # 添加性能日志
    driver.execute_cdp_cmd('Performance.enable', {})
    # 清空之前的日志
    driver.get_log('performance')


# 数据处理
def process_response():
    global remaining_balance
    all_records = []
    processed_ids = set()  # 去除重复内容
    try:
        monitor()
        captured_responses = []

        # # 配置CDP监听
        # driver.execute_cdp_cmd('Network.enable', {})

        # 清空之前的日志
        driver.get_log('performance')

        # 等待筛选条件生效
        time.sleep(3)

        # 获取总页数
        try:
            # 等待页码元素加载完成
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".finder-ui-desktop-pagination__num"))
            )
            page_elements = driver.find_elements(By.CSS_SELECTOR, ".finder-ui-desktop-pagination__num")

            if page_elements:
                # 获取当前显示的最后一个页码
                total_pages = int(page_elements[-1].text)
                # print(f"总页数: {total_pages}")
            else:
                total_pages = 1
        except TimeoutException:
            # print("只有一页数据")
            total_pages = 1
        except Exception as e:
            print(f"获取总页数时出错: {e}")
            total_pages = 1

        # 遍历每一页
        current_page = 1
        while current_page <= total_pages:
            # print(f"\n正在处理第 {current_page} 页...")

            # 如果不是第一页，点击对应页码
            if current_page > 1:
                try:
                    # 重新获取页码元素
                    page_elements = driver.find_elements(By.CSS_SELECTOR, ".finder-ui-desktop-pagination__num")
                    # 如果当前页码在可见页码中
                    page_found = False
                    for element in page_elements:
                        if element.text.strip() == str(current_page):
                            driver.execute_script("arguments[0].click();", element)
                            page_found = True
                            time.sleep(2)
                            break

                    # 如果当前页码不在可见范围内，需要处理
                    if not page_found:
                        # 找到并点击省略号后面的数字
                        ellipsis_elements = driver.find_elements(By.CSS_SELECTOR,
                                                                 ".finder-ui-desktop-pagination__ellipsis")
                        if ellipsis_elements:
                            # 点击省略号后的第一个数字
                            next_visible_page = driver.find_element(By.CSS_SELECTOR,
                                                                    ".finder-ui-desktop-pagination__num:last-child")
                            driver.execute_script("arguments[0].click();", next_visible_page)
                            time.sleep(2)
                            # 重试点击目标页码
                            page_elements = driver.find_elements(By.CSS_SELECTOR, ".finder-ui-desktop-pagination__num")
                            for element in page_elements:
                                if element.text.strip() == str(current_page):
                                    driver.execute_script("arguments[0].click();", element)
                                    time.sleep(2)
                                    break
                except Exception as e:
                    print(f"翻页失败: {e}")
                    break

            # 使用日志监听
            start_time = time.time()
            while time.time() - start_time < 5:  # 延长等待时间
                logs = driver.get_log('performance')
                for entry in logs:
                    try:
                        message = json.loads(entry['message'])['message']
                        if message['method'] == 'Network.responseReceived':
                            response = message['params']['response']
                            url = response.get('url', '')

                            if 'selectFeedPromotion' in url:
                                request_id = message['params']['requestId']
                                try:
                                    response_body = driver.execute_cdp_cmd('Network.getResponseBody', {
                                        'requestId': request_id
                                    })
                                    if response_body and 'body' in response_body:
                                        captured_responses.append(json.loads(response_body['body']))
                                except Exception as e:
                                    print(f"获取响应体时出错: {e}")
                    except Exception as e:
                        print(f"日志解析错误: {str(e)[:50]}")
                time.sleep(1.5)

            # 处理捕获的响应
            for response in captured_responses:
                if 'data' in response and 'orders' in response['data']:
                    for order in response['data']['orders']:
                        if 'orderInfo' in order:
                            # 获取订单基本信息
                            order_info = order['orderInfo']
                            promotion_id = order_info.get('promotionId')

                            # 检查是否已处理过该订单
                            if promotion_id in processed_ids:
                                continue
                            processed_ids.add(promotion_id)

                            quota = order_info.get('quota')
                            cost = order_info.get('cost')
                            status = order_info.get('status')

                            if status == 2:
                                quota = float(quota) / 10 if quota else 0
                                cost = float(cost) / 10 if cost else 0
                                remaining = quota - cost
                                remaining_balance.append(remaining)  # 添加到全局列表
                                # 添加到记录
                                all_records.append({
                                    "订单编号": promotion_id,
                                    "下单微信豆": quota,
                                    "已消耗微信豆": cost,
                                    "剩余微信豆": remaining,
                                    "状态": "加热中"
                                })

            current_page += 1

        # 输出结果
        if not all_records:
            record = "null"
            print(record)
        else:
            print("\n加热中订单列表：")
            for record in all_records:
                print(record)

    except Exception as e:
        print(f"处理响应时出错: {e}")
    finally:
        # 关闭网络监听
        driver.execute_cdp_cmd('Network.disable', {})


# 账户名称
def name():
    global name_text,video_account
    try:
        driver.execute_script(
            'document.querySelectorAll(".finder-ui-desktop-menu__popup_sub_menu")[2].style.display = "flex";')
        element = driver.find_element(By.XPATH,
                                      '/html/body/div/div/div[1]/div[1]/div/div/ul/li[3]/div/ul/li[1]/a/span/span/span')
        driver.execute_script("arguments[0].click();", element)
    except:
        print("找不到账户设置菜单")
    time.sleep(5)

    try:
        # 使用JavaScript获取视频号信息
        video_account = driver.execute_script("""
            const elements = document.querySelectorAll("div[class*='account-info-label']");
            for (const el of elements) {
                if (el.textContent.includes('视频号')) {
                    return el.nextElementSibling ? el.nextElementSibling.textContent.trim() : '';
                }
            }
            return '';
        """)
        name_text = driver.execute_script("""
                    const elements = document.querySelectorAll("div[class*='account-info-label']");
                    for (const el of elements) {
                        if (el.textContent.includes('账号主体')) {
                            return el.nextElementSibling ? el.nextElementSibling.textContent.trim() : '';
                        }
                    }
                    return '';
                """)

        if video_account:
            combined_name = f"{name_text}-{video_account}"
            print(f"主体名称:{name_text}")
            print(f"视频号名称：{video_account}")
            # print(f"完整账户信息: {combined_name}")
            return combined_name
        else:
            print("未找到视频号信息")
            return name_text

    except Exception as e:
        print(f"获取视频号信息失败: {str(e)}")
        return name_text


# 微信豆明细
def wechat_coins():
    try:
        global name_text
        # 定义所有余额类型
        balance_types = [
            "通用余额",  # 初始页面，不需要切换
            "充赠余额",
            "直赠余额",
            "专用余额",
            "cps内购激励余额",
            "cps广告激励余额",
            "加热广告激励余额"
        ]

        # 点击菜单
        driver.execute_script(
            'document.querySelectorAll(".finder-ui-desktop-menu__popup_sub_menu")[3].style.display = "flex";')
        time.sleep(2)
        element = driver.find_element(By.XPATH,
                                      '/html/body/div/div/div[1]/div[1]/div/div/ul/li[4]/div/ul/li[2]/a/span/span/span')
        driver.execute_script("arguments[0].click();", element)
        time.sleep(5)

        # 等待iframe加载
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "iframe"))
        )

        try:
            # 切换到iframe
            driver.switch_to.frame(iframe)
            time.sleep(2)

            # 遍历所有余额类型
            for balance_type in balance_types:
                print(f"\n正在处理 {balance_type} 数据...")
                all_table_data = []

                # 如果不是第一个类型（通用余额），需要切换类型
                if balance_type != "通用余额":
                    # 等待下拉菜单元素加载完成
                    dropdown_menu = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".weui-desktop-dropdown-menu"))
                    )
                    driver.execute_script('arguments[0].style.display = "flex";', dropdown_menu)
                    time.sleep(2)

                    # 根据类型选择对应的选项
                    if balance_type == "充赠余额":
                        driver.find_element(By.CSS_SELECTOR, "li:nth-child(2) div:nth-child(1)").click()
                    elif balance_type == "直赠余额":
                        driver.find_element(By.CSS_SELECTOR, "li:nth-child(3) div:nth-child(1)").click()
                    elif balance_type == "专用余额":
                        driver.find_element(By.CSS_SELECTOR, "li:nth-child(4) div:nth-child(1)").click()
                    elif balance_type == "cps内购激励余额":
                        driver.find_element(By.CSS_SELECTOR, "li:nth-child(5) div:nth-child(1)").click()
                    elif balance_type == "cps广告激励余额":
                        driver.find_element(By.CSS_SELECTOR, "li:nth-child(6) div:nth-child(1)").click()
                    elif balance_type == "加热广告激励余额":
                        driver.find_element(By.CSS_SELECTOR, "li:nth-child(7) div:nth-child(1)").click()

                time.sleep(5)

                # 验证数据是否存在
                try:
                    table = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, ".weui-desktop-table__loading-content__slot"))
                    )
                    # 检查元素文本是否包含"暂无数据"
                    if "暂无数据" in table.text:
                        print(f"{balance_type} 暂无数据")
                        continue
                    rows = table.find_elements(By.TAG_NAME, "tr")
                    if not rows:
                        print(f"{balance_type} 暂无数据")
                        continue
                except Exception as e:
                    print(f"{balance_type} 暂无数据: {e}")
                    continue

                # 下载文件
                download_button = driver.find_element(By.XPATH, '//a[contains(text(),"下载明细")]')
                # 清空临时目录中的文件
                for f in glob.glob(os.path.join(temp_download_dir, "*")):
                    try:
                        os.remove(f)
                    except:
                        pass
                
                # 点击下载
                download_button.click()

                # 等待文件下载完成
                wait_time = 0
                downloaded_file = None
                while wait_time < 20:  # 最多等待20秒
                    time.sleep(1)
                    wait_time += 1
                    files = glob.glob(os.path.join(temp_download_dir, "*.xlsx"))
                    if files:
                        # 检查文件是否完成下载(不再有.crdownload或.tmp扩展名)
                        incomplete_files = glob.glob(os.path.join(temp_download_dir, "*.crdownload")) + \
                                           glob.glob(os.path.join(temp_download_dir, "*.tmp"))
                        if not incomplete_files and len(files) == 1:  # 确保只有一个文件且已下载完成
                            downloaded_file = files[0]
                            break

                # 如果找到了下载文件，就重命名并移动它
                if downloaded_file:
                    # 创建文件名：combined_name加类型加微信豆明细加时间戳
                    combined_name = f"{name_text}-{video_account}" if video_account else name_text
                    timestamp = int(time.time())
                    new_filename = f"{combined_name}_{balance_type}_微信豆明细_{timestamp}.xlsx"
                    # 替换文件名中的非法字符
                    new_filename = new_filename.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace(
                        '?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                    new_filepath = os.path.join(download_dir, new_filename)
                    
                    # 移动文件到主下载目录并重命名
                    shutil.move(downloaded_file, new_filepath)
                    print(f"文件已重命名并移动: {new_filepath}")
                    
                    # 处理文件
                    process_downloaded_file(balance_type, video_account, name_text, new_filepath)
                else:
                    print("下载失败或超时")

        except Exception as e:
            print(f"处理 {balance_type} 数据时出错: {e}")
        finally:
            driver.switch_to.default_content()

    except Exception as e:
        print(f"处理微信豆明细时出错: {str(e)}")
    finally:
        pass


# # 重命名最新下载的Excel文件
# def rename_latest_excel(combined_name, balance_type):
#     try:
#         # 查找最新下载的Excel文件
#         excel_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
#         if not excel_files:
#             print("未找到Excel文件")
#             return None
#
#         latest_file = max(excel_files, key=os.path.getmtime)
#
#         # 创建新文件名
#         new_filename = f"{combined_name}_{balance_type}_微信豆明细.xlsx"
#         # 替换文件名中的非法字符
#         new_filename = new_filename.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace(
#             '?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
#         new_filepath = os.path.join(download_dir, new_filename)
#
#         # 如果目标文件已存在，先删除
#         if os.path.exists(new_filepath):
#             os.remove(new_filepath)
#
#         # 重命名文件
#         os.rename(latest_file, new_filepath)
#         return new_filepath
#     except Exception as e:
#         print(f"重命名文件时出错: {e}")
#         return None


# 处理下载的Excel文件并输出到控制台
def process_downloaded_file(balance_type, video_account, name_text, excel_path):
    try:
        # 创建combined_name用于查找重命名后的文件
        combined_name = f"{name_text}-{video_account}" if video_account else name_text
        # 替换combined_name中的非法字符
        safe_combined_name = combined_name.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*',
                                                                                                          '_').replace(
            '?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')

        # 查找已重命名的文件
        excel_filename = f"{safe_combined_name}_{balance_type}_微信豆明细.xlsx"
        excel_path = os.path.join(download_dir, excel_filename)

        if not os.path.exists(excel_path):
            print(f"未找到重命名后的文件: {excel_path}")
            # 如果找不到重命名后的文件，尝试查找最新下载的Excel文件
            excel_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
            if not excel_files:
                # print("未找到任何Excel文件")
                return
            excel_path = max(excel_files, key=os.path.getmtime)


        print(f"处理文件: {excel_path}")

        # 读取Excel文件
        df = pd.read_excel(excel_path)

        # 如果DataFrame为空，直接返回
        if df.empty:
            print(f"{balance_type} 文件内容为空")
            return

        # 添加额外的列
        df.insert(0, '主体', name_text)
        df.insert(1, '视频号名称', video_account)
        df.insert(2, '微信豆类型', balance_type)
        wechatcoin = tabulate(df, headers='keys', tablefmt='grid', showindex=False)
        # 输出表格到控制台
        print("\n数据表格输出:")
        print(wechatcoin)

        # 保存处理后的文件
        output_dir = os.path.join(download_dir, "处理后的明细")
        os.makedirs(output_dir, exist_ok=True)

        # 创建输出文件名
        output_filename = excel_filename
        output_path = os.path.join(output_dir, output_filename)

        # 保存处理后的文件
        df.to_excel(output_path, index=False)
        print(f"文件已保存至: {output_path}")
        return df
    except Exception as e:
        print(f"处理Excel文件时出错: {e}")
        return None


def cleanup_temp_directory():
    try:
        if os.path.exists(temp_download_dir):
            # 先删除文件夹中的所有文件
            for file in glob.glob(os.path.join(temp_download_dir, "*")):
                try:
                    os.remove(file)
                except Exception as e:
                    print(f"删除临时文件 {file} 失败: {e}")

            # 然后删除文件夹
            try:
                os.rmdir(temp_download_dir)
                print(f"临时文件夹 {temp_download_dir} 已删除")
            except Exception as e:
                print(f"删除临时文件夹 {temp_download_dir} 失败: {e}")
    except Exception as e:
        print(f"清理临时文件夹时出错: {e}")

try:
    success = main_automation(driver, download_dir)


finally:
    cleanup_temp_directory()
    print("脚本执行结束。")
    driver.quit()


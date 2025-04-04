import time
import random
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementNotInteractableException, \
    StaleElementReferenceException
from openpyxl.styles import Alignment

# 初始化浏览器
options = webdriver.ChromeOptions()
options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/91.0.4472.124')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_experimental_option('useAutomationExtension', False)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"""
})


def switch_to_iframe():
    """切换到 iframe 框架"""
    try:
        iframe = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "iFrame1"))
        )
        driver.switch_to.frame(iframe)
        print("已切换到 iframe 框架")
    except TimeoutException as e:
        print("❌ 无法切换到 iframe 框架:", e)
        raise


def get_latest_period(driver):
    """获取页面上的最新期号，使用用户提供的最新 XPath"""
    try:
        period_element = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, "//*[@id='historyData']/tr[1]/td[1]"))
        )
        period_text = period_element.text.strip()
        print(f"获取到的最新期号原始文本: '{period_text}'")
        if period_text.isdigit():
            return int(period_text)
        else:
            print(f"获取到的期号不是数字: '{period_text}'")
            return None
    except Exception as e:
        print(f"获取最新期号失败: {e}")
        return None


def get_existing_max_period():
    """获取现有 Excel 文件中的最大期号"""
    if os.path.exists("大乐透历史数据.xlsx"):
        try:
            df = pd.read_excel("大乐透历史数据.xlsx")
            max_period = df["期号"].max() if not df.empty else None
            print(f"现有 Excel 文件中的最大期号: {max_period}")
            return max_period
        except Exception as e:
            print(f"读取现有文件失败: {e}")
            return None
    print("未找到现有 Excel 文件")
    return None


def append_to_excel(data):
    """将数据追加到 Excel，确保格式清晰，最新期号在顶部，前区与后区之间空一列"""
    if not data:
        return False
    try:
        df_new = pd.DataFrame(data,
                              columns=["期号", "红球1", "红球2", "红球3", "红球4", "红球5", "分隔", "蓝球1", "蓝球2"])
        df_new["期号"] = df_new["期号"].astype(int)
        df_new["分隔"] = ""

        if os.path.exists("大乐透历史数据.xlsx"):
            df_existing = pd.read_excel("大乐透历史数据.xlsx")
            if "分隔" not in df_existing.columns:
                df_existing.insert(6, "分隔", "")
            df_combined = pd.concat([df_new, df_existing]).drop_duplicates(subset=["期号"]).sort_values(by="期号",
                                                                                                        ascending=False)
        else:
            df_combined = df_new.sort_values(by="期号", ascending=False)

        with pd.ExcelWriter("大乐透历史数据.xlsx", engine='openpyxl') as writer:
            df_combined.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            column_widths = {
                "期号": 12,
                "红球1": 10, "红球2": 10, "红球3": 10, "红球4": 10, "红球5": 10,
                "分隔": 5,
                "蓝球1": 10, "蓝球2": 10
            }
            for col_name, width in column_widths.items():
                col_idx = df_combined.columns.get_loc(col_name) + 1
                worksheet.column_dimensions[chr(65 + col_idx - 1)].width = width
            for row in worksheet.rows:
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        return True
    except Exception as e:
        print(f"保存数据失败: {e}")
        return False


try:
    driver.get("https://www.lottery.gov.cn/kj/kjlb.html?dlt")
    print("页面已加载")
    time.sleep(3)

    switch_to_iframe()
    latest_period = get_latest_period(driver)
    if latest_period is None:
        raise Exception("无法获取最新期号")

    existing_max = get_existing_max_period()
    start_period = existing_max + 1 if existing_max is not None else 7001  # 修改为 07001，去掉 20 前缀

    print(f"检测到最新期号: {latest_period}")
    print(f"现有数据最新期号: {existing_max if existing_max is not None else '无'}")
    print(f"计划从 {start_period} 期开始追加")

    # 添加额外调试信息
    print(f"start_period 类型: {type(start_period)}, 值: {start_period}")
    print(f"latest_period 类型: {type(latest_period)}, 值: {latest_period}")
    print(f"比较结果: start_period > latest_period = {start_period > latest_period}")

    # 条件判断逻辑
    if start_period > latest_period:
        print(f"条件判断: {start_period} > {latest_period}，没有新数据需要追加")
        driver.quit()
        exit()
    else:
        print(f"条件判断: {start_period} <= {latest_period}，开始抓取数据")

    data = []
    seen_periods = set()
    max_records = 10000
    page = 1

    while len(data) < max_records:
        print(f"正在爬取第 {page} 页...")
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//table//tr"))
            )
            table = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "table"))
            )
            rows = table.find_elements(By.TAG_NAME, "tr")
            print(f"找到 {len(rows)} 行数据")

            for _ in range(3):  # 重试3次以处理陈旧元素
                try:
                    for row in rows[2:]:  # 跳过表头
                        cols = row.find_elements(By.TAG_NAME, "td")
                        if len(cols) < 9:
                            continue

                        period = cols[0].text.strip()
                        if not period.isdigit():
                            continue

                        period_int = int(period)
                        if period_int < start_period or period_int > latest_period or period in seen_periods:
                            continue
                        seen_periods.add(period)

                        red_balls = []
                        for i in range(2, 7):  # 前区红球 td[2] 到 td[6]
                            red_ball = cols[i].text.strip()
                            if red_ball.isdigit():
                                red_balls.append(red_ball)
                            else:
                                break
                        if len(red_balls) != 5:
                            continue

                        blue_balls = []
                        for i in range(7, 9):  # 后区蓝球 td[7] 到 td[8]
                            blue_ball = cols[i].text.strip()
                            if "<span>" in cols[i].get_attribute('outerHTML'):
                                blue_ball = cols[i].find_element(By.TAG_NAME, "span").text.strip()
                            if blue_ball.isdigit():
                                blue_balls.append(blue_ball)
                            else:
                                break
                        if len(blue_balls) != 2:
                            continue

                        data.append([period] + red_balls + [""] + blue_balls)  # 插入空列
                    break
                except StaleElementReferenceException:
                    print("遇到陈旧元素引用，重新定位表格...")
                    table = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, "table"))
                    )
                    rows = table.find_elements(By.TAG_NAME, "tr")

            print(f"当前已抓取 {len(data)} 条记录")

            if data and int(data[-1][0]) <= start_period:
                print("已到达目标起始期号，停止抓取")
                break

            try:
                next_page = page + 1
                page_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//li[@onclick=\"kjCommonFun.goNextPage({next_page})\"]"))
                )
                page_button.click()
                print(f"成功翻页到第 {next_page} 页")
                time.sleep(random.uniform(3, 5))
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    f"//li[@class='number active' and @onclick=\"kjCommonFun.goNextPage({next_page})\"]"))
                )
                page += 1
            except TimeoutException:
                try:
                    next_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), '下一页')]"))
                    )
                    if "disabled" in next_button.get_attribute("class"):
                        print("已到最后一页")
                        break
                    next_button.click()
                    print("成功翻页到下一页")
                    time.sleep(random.uniform(3, 5))
                    WebDriverWait(driver, 10).until(EC.staleness_of(next_button))
                    page += 1
                except TimeoutException:
                    print("无法翻页，停止抓取")
                    break

        except TimeoutException:
            print("表格加载超时，跳过当前页")
            break

    if data:
        success = append_to_excel(data)
        print(f"✅ 成功追加 {len(data)} 条记录" if success else "❌ 数据保存失败")
    else:
        print("没有新数据需要追加")

except Exception as e:
    print(f"程序出错: {e}")
    print("当前页面源代码片段（前2000字符）:")
    print(driver.page_source[:2000])
finally:
    driver.quit()
    print("浏览器已关闭")
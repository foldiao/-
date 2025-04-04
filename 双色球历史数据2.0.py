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


def get_latest_period(driver):
    """获取页面上的最新期号"""
    try:
        period_element = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, "//table/tbody/tr[1]/td[1]"))
        )
        period_text = period_element.text.strip()
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
    if os.path.exists("双色球历史数据.xlsx"):
        try:
            df = pd.read_excel("双色球历史数据.xlsx")
            return df["期号"].max() if not df.empty else None
        except Exception as e:
            print(f"读取现有文件失败: {e}")
    return None


def reset_query_page(retry_count=3):
    """重置查询页面状态"""
    for attempt in range(retry_count):
        try:
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(2)
            custom_query_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//strong[@class='N-t' and text()='自定义查询']"))
            )
            custom_query_button.click()
            print("已点击‘自定义查询’")

            by_period_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='tj0' and text()='按期号']"))
            )
            by_period_button.click()
            print("已点击‘按期号’")
            return True
        except Exception as e:
            print(f"第 {attempt + 1} 次重置失败: {e}")
            if attempt < retry_count - 1:
                print("刷新页面并重试...")
                driver.refresh()
                time.sleep(random.uniform(2, 4))
            else:
                raise


def query_period_range(start_period, end_period, retry_count=3):
    """查询指定期号范围"""
    for attempt in range(retry_count):
        try:
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(2)

            start_input = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input.stcount"))
            )
            end_input = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input.endcount"))
            )

            driver.execute_script("arguments[0].value = '';", start_input)
            start_input.send_keys(str(start_period))
            print(f"已输入起始期号: {start_period}")
            time.sleep(random.uniform(1, 2))

            driver.execute_script("arguments[0].value = '';", end_input)
            end_input.send_keys(str(end_period))
            print(f"已输入结束期号: {end_period}")
            time.sleep(random.uniform(1, 2))

            start_query_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/div[2]/div[3]/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[2]/div"))
            )
            print(
                f"‘开始查询’按钮状态: 显示={start_query_button.is_displayed()}, 启用={start_query_button.is_enabled()}")
            driver.execute_script("arguments[0].click();", start_query_button)
            print("已通过 JavaScript 点击‘开始查询’")

            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, "//table//tr"))
            )
            print("查询结果已加载")
            return True
        except Exception as e:
            print(f"第 {attempt + 1} 次查询失败: {e}")
            if attempt < retry_count - 1:
                reset_query_page()
                time.sleep(random.uniform(2, 4))
            else:
                raise


def append_to_excel(data):
    """将数据追加到 Excel，确保格式清晰，最新期号在顶部，红球与蓝球之间空一列"""
    if not data:
        return False
    try:
        # 添加空列作为分隔
        df_new = pd.DataFrame(data,
                              columns=["期号", "红球1", "红球2", "红球3", "红球4", "红球5", "红球6", "分隔", "蓝球"])
        df_new["期号"] = df_new["期号"].astype(int)
        df_new["分隔"] = ""  # 空列作为红球和蓝球的分隔

        if os.path.exists("双色球历史数据.xlsx"):
            df_existing = pd.read_excel("双色球历史数据.xlsx")
            # 如果现有文件没有“分隔”列，添加空列
            if "分隔" not in df_existing.columns:
                df_existing.insert(7, "分隔", "")
            df_combined = pd.concat([df_new, df_existing]).drop_duplicates(subset=["期号"]).sort_values(by="期号",
                                                                                                        ascending=False)
        else:
            df_combined = df_new.sort_values(by="期号", ascending=False)

        # 保存到 Excel，设置宽松的列宽和居中对齐
        with pd.ExcelWriter("双色球历史数据.xlsx", engine='openpyxl') as writer:
            df_combined.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']

            # 设置列宽
            column_widths = {
                "期号": 12,
                "红球1": 10, "红球2": 10, "红球3": 10, "红球4": 10, "红球5": 10, "红球6": 10,
                "分隔": 5,  # 空列较窄
                "蓝球": 10
            }
            for col_name, width in column_widths.items():
                col_idx = df_combined.columns.get_loc(col_name) + 1
                worksheet.column_dimensions[chr(65 + col_idx - 1)].width = width

            # 设置居中对齐
            for row in worksheet.rows:
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        return True
    except Exception as e:
        print(f"保存数据失败: {e}")
        return False


try:
    driver.get("https://www.zhcw.com/kjxx/ssq/")
    print("页面已加载")
    time.sleep(3)

    latest_period = get_latest_period(driver)
    if latest_period is None:
        raise Exception("无法获取最新期号")

    existing_max = get_existing_max_period()
    start_period = existing_max + 1 if existing_max else 2003001

    if start_period > latest_period:
        print("没有新数据需要追加")
        driver.quit()
        exit()

    print(f"检测到最新期号: {latest_period}")
    print(f"现有数据最新期号: {existing_max if existing_max else '无'}")
    print(f"将从 {start_period} 期开始追加")

    reset_query_page()
    query_period_range(start_period, latest_period)

    data = []
    current_page = 1
    max_records = 10000

    while len(data) < max_records:
        print(f"正在爬取第 {current_page} 页...")
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//table//tr"))
            )
            table = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "table"))
            )
            print(f"表格内容（前500字符）: {table.text[:500]}")

            rows = table.find_elements(By.TAG_NAME, "tr")
            print(f"找到 {len(rows)} 行数据")

            for _ in range(3):  # 重试3次以处理陈旧元素
                try:
                    for row in rows:
                        cols = row.find_elements(By.TAG_NAME, "td")
                        if len(cols) < 4:
                            continue

                        period = cols[0].text.strip()
                        if not period.startswith("20"):
                            continue

                        try:
                            period_int = int(period)
                            if period_int < start_period or period_int > latest_period:
                                continue
                        except ValueError:
                            continue

                        red_balls = []
                        blue_ball = ""

                        red_spans = cols[2].find_elements(By.CLASS_NAME, "jqh")
                        if red_spans:
                            red_balls = [span.text.strip() for span in red_spans]
                        else:
                            red_text = cols[2].text.strip().replace("\n", "")
                            red_balls = [red_text[i:i + 2] for i in range(0, len(red_text), 2) if
                                         i + 2 <= len(red_text)]

                        blue_ball = cols[3].text.strip()

                        if len(red_balls) == 6 and blue_ball.isdigit():
                            data.append([period] + red_balls + [""] + [blue_ball])  # 插入空列
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
                current_page += 1
                page_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//a[@title='{current_page}']"))
                )
                page_button.click()
                print(f"成功翻页到第 {current_page} 页")
                time.sleep(random.uniform(3, 5))
                WebDriverWait(driver, 10).until(EC.staleness_of(page_button))
            except TimeoutException:
                try:
                    next_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '下一页')]"))
                    )
                    if "disabled" in next_button.get_attribute("class"):
                        print("已到最后一页")
                        break
                    next_button.click()
                    print("成功翻页到下一页")
                    time.sleep(random.uniform(3, 5))
                    WebDriverWait(driver, 10).until(EC.staleness_of(next_button))
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
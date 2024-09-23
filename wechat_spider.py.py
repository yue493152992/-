import time
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from datetime import datetime

# 设置搜索关键字和爬取页数
keyword = "AI"
pages_to_scrape = 5

# 设置 WebDriver 路径
chrome_driver_path = r"D:\chromedriver\chromedriver.exe" # 替换为您自己的 chromedriver 路径

# 设置 ChromeDriver 服务
service = Service(chrome_driver_path)

# 初始化 WebDriver
options = webdriver.ChromeOptions()

driver = webdriver.Chrome(service=service, options=options)

# 打开网页
driver.get("https://weixin.sogou.com/")

# 输入关键字并点击搜索按钮
search_box = driver.find_element(By.ID, "query")
search_box.send_keys(keyword)
search_button = driver.find_element(By.XPATH, '//input[@value="搜文章"]')
search_button.click()

# 创建 Excel 文件
wb = Workbook()
ws = wb.active
ws.append(["标题", "摘要", "链接", "来源"])

# 爬取内容
for page in range(pages_to_scrape):
    print(f"正在爬取第 {page + 1} 页...")
    # 等待页面加载完成
    time.sleep(5)
    # 获取搜索结果列表
    search_results = driver.find_elements(By.XPATH, '//ul[@class="news-list"]/li')
    # 提取信息并写入 Excel
    for result in search_results:
        title = result.find_element(By.XPATH, './/h3/a').text
        abstract = result.find_element(By.XPATH, './/p[@class="txt-info"]').text
        link = result.find_element(By.XPATH, './/h3/a').get_attribute("href")
        source = result.find_element(By.XPATH, './/div[@class="s-p"]').text
        # 去除日期部分
        date_index = source.find("20")  # 找到日期的起始位置
        if date_index != -1:
            source = source[:date_index]  # 截取日期之前的内容
        ws.append([title, abstract, link, source])
    # 滚动到页面底部
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    # 等待下一页按钮出现
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'sogou_next')))
    except TimeoutException:
        print("已到达最后一页")
        break
    # 点击下一页按钮
    next_page_button = driver.find_element(By.ID, 'sogou_next')
    next_page_button.click()

# 保存 Excel 文件
current_time = datetime.now().strftime("%Y%m%d%H%M%S")
file_name = f"{keyword}_微信_{current_time}.xlsx"
wb.save(file_name)

print(f"爬取完成，结果保存在 {file_name}")

# 关闭浏览器
driver.quit()